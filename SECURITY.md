# Security Policy

## Supported Versions

| Version | Supported |
| ------- | --------- |
| 2.0.x   | yes       |
| < 2.0   | no        |

## Reporting a Vulnerability

Please **don't** open a public issue for security problems.

Use GitHub's [private vulnerability reporting](https://github.com/IOST-ASCOL/doc2tex/security/advisories/new)
instead and include steps to reproduce (a sample file helps a lot). We'll get back to you as soon
as we can and credit you in the release notes if you'd like.

A few things to keep in mind:

- The web UI (`doc2tex-web`) is meant for your own computer or a trusted network. It only
  listens on `127.0.0.1` unless you set `DOC2TEX_HOST`.
- Never run it with `DOC2TEX_DEBUG=1` on a network other people can reach - the Flask
  debugger lets anyone who can open the page run code on your machine.
- `--pdf` runs `latexmk`/`pdflatex` on the converted file. Only compile documents you trust.
