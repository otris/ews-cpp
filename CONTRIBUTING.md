# How to contribute

We like to encourage you to contribute to the repository.  This should be as
easy as possible for you but there are a few things to consider when
contributing.  The following guidelines for contribution should be followed if
you want to submit a pull request.

## How to prepare

* You need a [GitHub account](https://github.com/signup/free)
* Submit an [issue ticket](https://github.com/otris/ews-cpp/issues) for your
  issue if there isn't one yet.
  * Describe the issue and include steps to reproduce if it's a bug.
  * Ensure to mention the earliest version that you know is affected.
* If you are able and want to fix this, fork the repository on GitHub.

## Make Changes

* In your forked repository, create a topic branch for your upcoming patch.
  (e.g. `feature-autodiscover` or `bugfix-windows-crash`)
  * Usually this is based on the master branch.
  * Create a branch based on master; `git branch
  my-work master` then checkout the new branch with `git
  checkout my-work`.  Please avoid working directly on the `master` branch.
* Make sure you stick to the existing coding style.
* Prior to committing, stage your changes and run `git clang-format` if it is
  available on your platform.
* Make commits of logical units and describe them properly. See
  [How to Write a Git Commit Message](http://chris.beams.io/posts/git-commit/).
* Check for unnecessary white space with `git diff --check` before committing.
* If possible, submit tests to your patch / new feature so it can be tested easily.
* Assure nothing is broken by running all the tests, if possible.
* You should use `git pull --rebase` instead of `git pull` to avoid generating
  a non-linear history in your clone. To configure `git pull` to pass
  `--rebase` by default on the master branch, run the following command:

 ```
 git config branch.master.rebase true
 ```

## Submit Changes

* Push your changes to a topic branch in your fork of the repository.
* Open a pull request to the original repository and choose the right original
  branch you want to patch.
* If not done in commit messages (which you really should do) please reference
  and update your issue with the code changes. But _please do not close the
  issue yourself_.
* Even if you have write access to the repository, do not directly push or
  merge pull-requests. Let another team member review your pull request and
  approve.

# Additional Resources

* [General GitHub documentation](http://help.github.com/)
* [GitHub pull request documentation](http://help.github.com/send-pull-requests/)
* [Read the Issue Guidelines by @necolas](https://github.com/necolas/issue-guidelines/blob/master/CONTRIBUTING.md) for more details

