import sys

from md2pptx.main import cli

if __name__ == "__main__" and not __package__:
    # This should never happen when installed from pip.
    # This workaround is NOT bulletproof, rather brittle as many edge
    # cases are not covered
    # See http://stackoverflow.com/a/28154841/2479038

    print(
        "warning: running package directly, risking ImportError",
        file=sys.stderr,
    )


if __name__ == "__main__":
    # Execute only if run as a script
    # See https://docs.python.org/3/library/__main__.html
    cli()
