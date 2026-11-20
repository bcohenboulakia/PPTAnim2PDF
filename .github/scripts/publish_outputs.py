"""Publish explicit output paths without merging source branches or rewriting history."""

import argparse
import json
from pathlib import Path, PurePosixPath
import shutil
import subprocess
import sys
import tempfile
import uuid


def git(*args, cwd=None, check=True):
    return subprocess.run(
        ["git", *map(str, args)], cwd=cwd, text=True,
        stdout=subprocess.PIPE, stderr=subprocess.PIPE, encoding="utf-8", check=check,
    )


def relative_path(name):
    path = PurePosixPath(name)
    if (not name or path.is_absolute() or any(part in ("..", ".") for part in path.parts)
            or any(char in name for char in ("\\", ":", "\n", "\r")) or str(path) != name):
        raise ValueError(f"Invalid repository-relative path: {name!r}")
    return path


def remote_head(branch, repo="."):
    # refs/heads/<branch> is fetched explicitly. No local branch checkout/merge.
    exists = git("ls-remote", "--exit-code", "origin", f"refs/heads/{branch}", cwd=repo, check=False)
    if exists.returncode == 2:
        return None
    if exists.returncode != 0:
        raise RuntimeError(f"Cannot check remote {branch}: {exists.stderr.strip()}")
    fetch = git("fetch", "--no-tags", "origin", f"refs/heads/{branch}", cwd=repo)
    return git("rev-parse", "FETCH_HEAD", cwd=repo).stdout.strip()


def overlap_with_remote(repo, prior_head, current_head, paths):
    if prior_head == current_head:
        return []
    if prior_head is None:
        # Another tool may have just created Students from main. Content
        # copied unchanged from main is not a competing publication.
        collisions = []
        for p in paths:
            remote_blob = git("rev-parse", "--verify", f"{current_head}:{p}",
                              cwd=repo, check=False)
            if remote_blob.returncode != 0:
                continue
            source_blob = git("rev-parse", "--verify", f"HEAD:{p}",
                              cwd=repo, check=False)
            if source_blob.returncode != 0 or source_blob.stdout.strip() != remote_blob.stdout.strip():
                collisions.append(p)
        return collisions
    changed = git("diff", "--name-only", "-z", prior_head, current_head, "--", *sorted(paths), cwd=repo)
    return [p for p in changed.stdout.split("\0") if p]


def _push(repo, worktree, branch):
    return git("push", "origin", f"HEAD:refs/heads/{branch}", cwd=worktree, check=False)


def publish_files(additions, removals=(), repo=".", branch="Students", max_attempts=4, message="Update generated materials"):
    if (not branch or branch == "main" or
            git("check-ref-format", "--branch", branch, cwd=repo, check=False).returncode != 0):
        raise ValueError(f"Invalid output branch: {branch!r}")
    repo = str(Path(repo).resolve())
    additions = {str(relative_path(name)): Path(path).resolve() for name, path in additions.items()}
    removals = {str(relative_path(name)) for name in removals} - additions.keys()
    for name in set(additions) | removals:
        if PurePosixPath(name).parts[0].lower() == ".git":
            raise ValueError(f"Cannot publish Git metadata: {name}")
    for path in additions.values():
        if not path.is_file():
            raise FileNotFoundError(path)
    owned_paths = set(additions) | removals
    initial_head = None

    for attempt in range(1, max_attempts + 1):
        head = remote_head(branch, repo)
        if attempt == 1:
            initial_head = head
        else:
            if head is None:
                raise RuntimeError("Output branch disappeared during publication")
            collisions = overlap_with_remote(repo, initial_head, head, owned_paths)
            if collisions:
                raise RuntimeError("Concurrent changes to publication paths: " + ", ".join(collisions))
        with tempfile.TemporaryDirectory(prefix="nbworkshop-publish-") as tmp:
            worktree = str((Path(tmp) / "worktree").resolve())
            if Path(worktree).parent != Path(tmp).resolve():
                raise RuntimeError("Unexpected temporary worktree path")
            orphan_branch = None
            git("worktree", "add", "--detach", worktree, head or "HEAD", cwd=repo)
            try:
                if head is None:
                    orphan_branch = "nbworkshop-empty-" + uuid.uuid4().hex[:12]
                    git("checkout", "--orphan", orphan_branch, cwd=worktree)
                    git("rm", "-rf", "--ignore-unmatch", "--", ".", cwd=worktree)
                    git("clean", "-fdx", "-q", cwd=worktree)
                for path in sorted(removals):
                    target = Path(worktree, path)
                    if target.is_file() or target.is_symlink():
                        target.unlink()
                for name, src in sorted(additions.items()):
                    dest = Path(worktree, name)
                    dest.parent.mkdir(parents=True, exist_ok=True)
                    shutil.copyfile(src, dest)
                git("add", "-A", cwd=worktree)
                if git("diff", "--cached", "--quiet", cwd=worktree, check=False).returncode == 0:
                    return {"status": "unchanged", "attempts": attempt}
                git("-c", "user.name=github-actions[bot]",
                    "-c", "user.email=github-actions[bot]@users.noreply.github.com",
                    "commit", "-m", message, cwd=worktree)
                push = _push(repo, worktree, branch)
                if push.returncode == 0:
                    return {"status": "published", "attempts": attempt}
                new_head = remote_head(branch, repo)
                if new_head == head:
                    raise RuntimeError("Git push failed: " + push.stderr.strip())
                print(f"Students advanced during push; retry {attempt}/{max_attempts}", file=sys.stderr)
            finally:
                git("worktree", "remove", "--force", worktree, cwd=repo, check=False)
                if orphan_branch:
                    git("branch", "-D", orphan_branch, cwd=repo, check=False)
    raise RuntimeError(f"Could not publish to {branch} after {max_attempts} attempts")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--file", action="append", default=[])
    parser.add_argument("--remove", action="append", default=[])
    parser.add_argument("--branch", default="Students")
    parser.add_argument("--message", default="Update converted presentation")
    args = parser.parse_args()
    if not args.file and not args.remove:
        parser.error("At least one --file or --remove is required")
    additions = {str(relative_path(name)): Path(name) for name in args.file}
    print(json.dumps(publish_files(additions, args.remove, branch=args.branch, message=args.message)))


if __name__ == "__main__":
    main()
