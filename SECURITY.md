# Security policy

VsOfx opens documents that often come from other people, so parsing bugs are security bugs.
Its design goal is that a document can never reach the network, run code, or touch files other
than the one being edited (see the README section "Security: no network, enforced").

## Supported versions

Only the latest release on the Marketplace / GitHub Releases gets fixes.

## Reporting a vulnerability

Please **do not open a public issue**. Use GitHub's private reporting instead:
**Security → Report a vulnerability** on <https://github.com/akharon/VsOfx/security/advisories/new>.

Include the VS Code version, the VsOfx version, and if possible a file that shows the problem.
You can expect a first answer within a week. Once a fix is released, the advisory is published
and you are credited unless you ask not to be.

In scope, for example: a document that makes the webview load a remote resource, run script,
escape the webview sandbox, open a non-web link, or corrupt/overwrite files other than itself.
