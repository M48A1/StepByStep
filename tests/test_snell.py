"""Run with python3 -m unittest discover -s tests -v. No root or network needed."""
import os
from pathlib import Path
import shlex
import subprocess
import tempfile
import unittest


SOURCE = Path(__file__).resolve().parents[1] / "snell.sh"


class SnellTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix="snell-test-")
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.bin = self.root / "bin"
        self.bin.mkdir()
        self.units = self.root / "units"
        self.units.mkdir()
        self.mock = self.root / "mock"
        self.mock.mkdir()
        self.conf = self.root / "etc/snell"
        self.unit = self.units / "snell-server.service"
        self.manager = self.bin / "snell"
        self.backups = self.root / "backups"
        self.sysdir = self.root / "sysctl.d"
        self.sysdir.mkdir()
        self.sysfile = self.sysdir / "99-snell-tfo.conf"
        self.kernel = self.root / "kernel"
        self.kernel.write_text("1024\n")
        self.ipv6 = self.root / "if_inet6"
        self.ipv6.write_text("20010db8000000000000000000000001 02 40 00 80 eth0\n")
        text = SOURCE.read_text()
        for old, new in {
            "/usr/local/bin": self.bin,
            "/etc/snell": self.conf,
            "/etc/systemd/system": self.units,
            "/var/backups/snell": self.backups,
            "/etc/sysctl.d": self.sysdir,
            "/proc/net/if_inet6": self.ipv6,
        }.items():
            text = text.replace(old, str(new))
        self.script = self.root / "snell.sh"
        self.script.write_text(text)
        mocks = {
            "uname": "echo x86_64",
            "sleep": "exit 0",
            "curl": '''[[ ${FAIL_DOWNLOAD:-0} != 1 ]] || exit 22
while [[ $# -gt 0 ]]; do
  if [[ $1 == -o ]]; then touch "$2"; exit; fi
  shift
done''',
            "unzip": '''if [[ $1 == -p ]]; then printf '#!/bin/bash\nexit 0\n'; fi''',
            "ss": '''if [[ ${BUSY_PORT:-0} == 1 ]]; then
  echo 'tcp LISTEN 0 128 0.0.0.0:34567 0.0.0.0:*'
fi''',
            "sysctl": '''if [[ $1 == -n ]]; then
  if [[ $2 == net.ipv6.bindv6only ]]; then echo "${BINDV6ONLY:-0}"; else cat "$KERNEL"; fi
else
  value=${3#*=}
  [[ ${FAIL_TFO:-0} != 1 || $value == 1024 ]] || exit 1
  printf '%s\n' "$value" > "$KERNEL"
fi''',
            "systemctl": '''echo "$*" >> "$ROOT/events"
case "$1" in
  cat) [[ -f $UNIT ]] ;;
  is-active) [[ -f $ROOT/active ]] ;;
  is-enabled) [[ -f $ROOT/enabled ]] ;;
  stop) rm -f "$ROOT/active" ;;
  disable) rm -f "$ROOT/enabled" ;;
  enable)
    touch "$ROOT/enabled"
    if [[ ${2:-} == --now && ${FAIL_START:-0} != 1 ]]; then touch "$ROOT/active"; fi ;;
  start) touch "$ROOT/active" ;;
  *) exit 0 ;;
esac''',
        }
        for name, body in mocks.items():
            path = self.mock / name
            path.write_text("#!/bin/bash\n" + body + "\n")
            path.chmod(0o755)
        self.env = {
            **os.environ,
            "PATH": str(self.mock) + ":" + os.environ["PATH"],
            "ROOT": str(self.root), "UNIT": str(self.unit), "KERNEL": str(self.kernel),
            "SNELL_VERSION": "5.0.1", "SNELL_PORT": "34567",
            "SNELL_NETWORK": "ipv4", "SNELL_TFO": "false",
            "SNELL_REUSE": "true", "SNELL_MODE": "default",
            "SNELL_HOST": "proxy.example.com", "SNELL_PSK": "a" * 32,
        }

    def shell(self, code, success=True, **env):
        result = subprocess.run(
            ["bash", "-c", f"source {shlex.quote(str(self.script))}; " + code],
            env={**self.env, **env}, stdin=subprocess.DEVNULL,
            capture_output=True, text=True,
        )
        self.assertEqual(result.returncode == 0, success, result.stdout + result.stderr)
        return result

    def install(self, reinstall=False, success=True, **env):
        return self.shell(
            "collect_options() { :; }; install_dependencies() { :; }; "
            + "install_snell " + ("true" if reinstall else "false"), success, **env,
        )

    def files(self):
        return {
            str(p.relative_to(self.root)): (p.read_bytes(), p.stat().st_mode & 0o777)
            for parent in (self.conf, self.bin, self.units, self.sysdir)
            for p in parent.rglob("*") if p.is_file()
        }

    def test_install_and_manager_config(self):
        result = self.install()
        self.assertIn("ip-version=v4-only", result.stdout)
        self.assertIn("ipv6 = false", (self.conf / "snell-server.conf").read_text())
        self.assertEqual(self.manager.stat().st_mode & 0o777, 0o755)
        result = subprocess.run(["bash", str(self.manager), "--help"], capture_output=True, text=True)
        self.assertEqual(result.returncode, 0)
        self.assertIn("reinstall", result.stdout)
        result = self.shell("show_config")
        self.assertIn("psk = " + "a" * 32, result.stdout)
        self.assertIn("RestrictAddressFamilies=AF_INET AF_UNIX AF_NETLINK", result.stdout)

    def test_reinstall_success_and_persistent_backup(self):
        self.install()
        extra = self.conf / "custom-note"
        extra.write_text("keep this")
        before = (self.conf / "snell-server.conf").read_bytes()
        result = self.install(True, SNELL_VERSION="6.0.0", SNELL_NETWORK="ipv6",
                              SNELL_HOST="2001:db8::1", SNELL_MODE="unshaped", SNELL_TFO="true")
        config = (self.conf / "snell-server.conf").read_text()
        self.assertIn("dns-ip-preference = ipv6-only", config)
        self.assertIn("listen = [2001:0db8:0000:0000:0000:0000:0000:0001]:34567", config)
        self.assertIn("[2001:db8::1]", result.stdout)
        self.assertIn("ip-version=v6-only", result.stdout)
        self.assertIn("RestrictAddressFamilies=AF_INET6 AF_UNIX AF_NETLINK", self.unit.read_text())
        self.assertEqual(extra.read_text(), "keep this")
        backup = next(self.backups.iterdir())
        self.assertEqual((backup / "1/snell-server.conf").read_bytes(), before)
        self.assertEqual(backup.stat().st_mode & 0o777, 0o700)
        self.assertEqual(self.kernel.read_text().strip(), "1027")
        self.assertTrue((self.root / "active").exists())

    def test_failed_reinstall_restores_all_files_and_service(self):
        self.install()
        (self.conf / "custom-note").write_text("keep")
        self.manager.write_text("#!/bin/bash\n# snell-manager-managed\necho old-manager\n")
        before = self.files()
        self.install(True, False, FAIL_START="1", SNELL_VERSION="6.0.0", SNELL_TFO="true")
        self.assertEqual(self.files(), before)
        self.assertEqual(self.kernel.read_text().strip(), "1024")
        self.assertTrue((self.root / "active").exists())
        self.assertTrue((self.root / "enabled").exists())

    def test_failed_reinstall_preserves_stopped_disabled_state(self):
        self.install()
        (self.root / "active").unlink()
        (self.root / "enabled").unlink()
        before = self.files()
        self.install(True, False, FAIL_START="1")
        self.assertEqual(self.files(), before)
        self.assertFalse((self.root / "active").exists())
        self.assertFalse((self.root / "enabled").exists())

    def test_download_failure_does_not_stop_existing_service(self):
        self.install()
        before = self.files()
        (self.root / "events").write_text("")
        self.install(True, False, FAIL_DOWNLOAD="1")
        self.assertEqual(self.files(), before)
        self.assertNotIn("stop", (self.root / "events").read_text())
        self.assertTrue((self.root / "active").exists())

    def test_port_conflict_restores_running_installation(self):
        self.install()
        before = self.files()
        self.install(True, False, BUSY_PORT="1")
        self.assertEqual(self.files(), before)
        self.assertTrue((self.root / "active").exists())

    def test_initial_install_failure_cleans_manager_and_files(self):
        self.install(success=False, FAIL_START="1", SNELL_TFO="true")
        self.assertFalse(self.conf.exists())
        self.assertFalse(self.manager.exists())
        self.assertFalse(self.unit.exists())
        self.assertFalse(self.sysfile.exists())
        self.assertEqual(self.kernel.read_text().strip(), "1024")

    def test_existing_tfo_configuration_survives_rollback(self):
        self.install(SNELL_TFO="true")
        before = self.files()
        self.install(True, False, FAIL_START="1", SNELL_TFO="true")
        self.assertEqual(self.files(), before)
        self.assertEqual(self.kernel.read_text().strip(), "1027")

    def test_foreign_manager_not_overwritten(self):
        self.manager.write_text("do not replace")
        self.install(success=False)
        self.assertEqual(self.manager.read_text(), "do not replace")

    def test_no_ipv6_rejected_before_modifications(self):
        self.ipv6.write_text("")
        self.install(success=False, SNELL_VERSION="6.0.0", SNELL_NETWORK="ipv6")
        self.assertFalse(self.conf.exists())

    def test_network_modes_and_version_constraints(self):
        for version, network, expected in [
            ("5.0.1", "ipv4", "ipv6 = false"),
            ("5.0.1", "dual", "ipv6 = true"),
            ("6.0.0", "ipv4", "dns-ip-preference = ipv4-only"),
            ("6.0.0", "ipv6", "dns-ip-preference = ipv6-only"),
            ("6.0.0", "dual", "listen = 0.0.0.0:34567,[::]:34567"),
        ]:
            with self.subTest(version=version, network=network):
                result = self.shell("validate_options; prepare_network; write_server_config",
                                    SNELL_VERSION=version, SNELL_NETWORK=network)
                self.assertIn(expected, result.stdout)
        self.shell("validate_options", False, SNELL_NETWORK="ipv6")
        self.shell("validate_options; prepare_network", False, SNELL_NETWORK="dual", BINDV6ONLY="1")


if __name__ == "__main__":
    unittest.main()
