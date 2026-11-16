import pytest

from tests.watchdog import watch_session


def pytest_sessionstart(session: pytest.Session) -> None:
    """A guard that regresses must cost a failed run, not the machine's memory."""
    watch_session()


def pytest_addoption(parser: pytest.Parser) -> None:
    parser.addoption(
        "--tier",
        choices=["full", "slow"],
        default="full",
        help="'full' (default) or 'slow' (full, plus the costly cases, as CI runs it).",
    )


def pytest_collection_modifyitems(config: pytest.Config, items: list[pytest.Item]) -> None:
    if config.getoption("--tier") == "slow":
        return
    skip_slow = pytest.mark.skip(reason="costly: run with --tier slow")
    for item in items:
        if "slow" in item.keywords:
            item.add_marker(skip_slow)
