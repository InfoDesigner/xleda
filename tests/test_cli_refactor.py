from typer.testing import CliRunner

from xleda import cli as cli_module


runner = CliRunner()


def test_cli_commands_delegate_to_cli_class(monkeypatch):
    calls = []

    class DummyCLI:
        def install(self):
            calls.append("install")

        def uninstall(self):
            calls.append("uninstall")

        def version(self):
            calls.append("version")

        def wb(self, **kwargs):
            calls.append(("wb", kwargs["data"], kwargs.get("file_name")))

    monkeypatch.setattr(cli_module, "CLI", DummyCLI)

    assert runner.invoke(cli_module.cli, ["install"]).exit_code == 0
    assert runner.invoke(cli_module.cli, ["uninstall"]).exit_code == 0
    assert runner.invoke(cli_module.cli, ["version"]).exit_code == 0
    assert runner.invoke(cli_module.cli, ["wb", "sample.csv"]).exit_code == 0

    assert calls == [
        "install",
        "uninstall",
        "version",
        ("wb", "sample.csv", None),
    ]
