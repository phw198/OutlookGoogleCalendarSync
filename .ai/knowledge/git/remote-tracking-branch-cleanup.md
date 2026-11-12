---
type: How-to
title: "Removing a local remote-tracking branch"
description: "Remove a remote-tracking reference from the local repository without deleting its branch on the remote."
resource: https://git-scm.com/docs/git-branch
tags:
  - Git
  - branch
  - remote-tracking
sources:
  - resource: https://git-scm.com/docs/git-branch
    title: "git-branch documentation"
  - resource: https://git-scm.com/docs/git-fetch
    title: "git-fetch documentation"
---

# Summary

A remote-tracking branch such as `ljlongwing/bugfix/issue-2264/organiser-not-self` is a local reference to a branch advertised by a remote. Deleting that reference locally does not delete the branch on the remote.

# Remove one local reference

Use `git branch -r -d <remote>/<branch>` to delete one remote-tracking reference:

```shell
git branch -r -d ljlongwing/bugfix/issue-2264/organiser-not-self
```

If the branch still exists on the remote, a later fetch can recreate the local reference.

# Prune stale references

If the remote branch has already been deleted and stale references need to be removed, use:

```shell
git fetch --prune <remote>
```

Pruning removes stale local remote-tracking references for that remote; it does not delete branches on the remote.

# Restrict which remote branches are fetched

To add one branch from the `ljlongwing` remote to the branches fetched on future fetches, without changing the fetch configuration for `origin`, use:

```shell
git remote set-branches --add ljlongwing bugfix/issue-2264/organiser-not-self
git fetch ljlongwing
```

This changes only the branch list for the `ljlongwing` remote; it does not change `origin` or any branches on the remote server. Confirm the configured refspecs with:

```shell
git config --get-all remote.ljlongwing.fetch
```

`--add` preserves the `ljlongwing` remote's existing fetch refspecs. If that remote already has a wildcard refspec, it will continue to fetch all its branches; adding one branch does not narrow that wildcard. To restrict only `ljlongwing` to this branch, replace that remote's branch list with:

```shell
git remote set-branches ljlongwing bugfix/issue-2264/organiser-not-self
git fetch --prune ljlongwing
```

This still leaves `origin`'s fetch configuration unchanged. Narrowing a refspec does not necessarily remove remote-tracking refs created by earlier broad fetches. Remove unwanted existing local refs explicitly with `git branch -r -d ljlongwing/<branch>`; do not assume pruning will delete refs that are no longer covered by the configured refspec.

# Do not confuse local cleanup with remote deletion

`git push <remote> --delete <branch>` deletes the branch from the remote. Use it only when the user explicitly requests remote branch deletion. If `git branch -r -d` refuses, inspect Git's reason rather than escalating to force deletion automatically.
