"""请求端校验器 _validate_endpoint 的回归钉。

_validate_endpoint（ai_assist.py）是请求侧的 SSRF 风格守卫：http/https-only、
缺主机名拒收、环回与 .localhost 放行（DESIGN §14.3 的本地 Ollama）、其余私网
地址默认拒收、allow_lan_endpoint 显式 opt-in 放行；域名按解析出的地址逐一判定。
此前全库零覆盖——删掉它全量 pytest 仍绿。这里逐分支直测，重构时守卫不再能
静默失效。

域名分支一律 monkeypatch socket.getaddrinfo：不做真实 DNS（CI 上解析结果
不可控，也无法构造「解析到私网」的确定域名）；测试只调校验器，从不发请求。
"""

import socket

import pytest

from conftest import ROOT

TOOLS = ROOT / "tools"

import sys
sys.path.insert(0, str(TOOLS))

import ai_assist  # noqa: E402


def _resolved(host, *addrs):
    """getaddrinfo 替身：把 host「解析」到给定地址（形同真实返回的裁剪版）。"""
    def fake(host_, port, *args, **kwargs):
        assert host_ == host, f"意外解析请求：{host_}"
        return [(socket.AF_INET, 0, 0, "", (a, 0)) for a in addrs]
    return fake


# 环回/私网地址均为 RFC 1918 / RFC 4291 文档段示例，不是可用凭据。
@pytest.mark.parametrize("url", [
    "file:///C:/Windows/system32/config",
    "ftp://example.invalid/v1",
    "gopher://example.invalid",
])
def test_non_http_schemes_are_rejected(url):
    """urllib 对 file:// 等协议照发不误，校验器必须先拦 scheme。"""
    with pytest.raises(ValueError, match="只允许 http/https"):
        ai_assist._validate_endpoint({}, url)


def test_missing_hostname_is_rejected():
    """scheme 合法但解析不出主机名（http:///v1）→ 拒收。"""
    with pytest.raises(ValueError, match="缺主机名"):
        ai_assist._validate_endpoint({}, "http:///v1")


@pytest.mark.parametrize("addr", [
    "192.168.1.5", "10.0.0.2", "172.16.0.9", "169.254.1.1",  # IPv4 私网/链路本地
    "[fc00::1]", "[fe80::1]",                                # IPv6 ULA/链路本地（URL 须带方括号）
])
def test_non_global_addresses_are_rejected_by_default(addr):
    """非全局地址的字面量默认拒收，报错文案给出 opt-in 开关名。"""
    with pytest.raises(ValueError, match="allow_lan_endpoint"):
        ai_assist._validate_endpoint({}, f"http://{addr}/v1")


def test_hostname_resolving_to_private_is_rejected(monkeypatch):
    """域名字面量常规，但解析出私网地址 → 仍拒（F2-4 的核心场景）。"""
    monkeypatch.setattr(ai_assist.socket, "getaddrinfo",
                        _resolved("intranet.example.invalid", "10.0.0.7"))
    with pytest.raises(ValueError, match="allow_lan_endpoint"):
        ai_assist._validate_endpoint({}, "http://intranet.example.invalid/v1")


def test_any_resolved_address_failing_rejects(monkeypatch):
    """解析出多个地址时逐一判定：一个私网即整单拒收。"""
    monkeypatch.setattr(
        ai_assist.socket, "getaddrinfo",
        _resolved("dual.example.invalid", "93.184.216.34", "192.168.0.9"))
    with pytest.raises(ValueError, match="allow_lan_endpoint"):
        ai_assist._validate_endpoint({}, "http://dual.example.invalid/v1")


def test_opt_in_allows_private_endpoint():
    """allow_lan_endpoint=True 显式放行内网模型服务。"""
    ai_assist._validate_endpoint({"allow_lan_endpoint": True},
                                 "http://192.168.1.5:11434/v1")


def test_loopback_and_localhost_are_allowed():
    """环回地址与 localhost/.localhost 放行——DESIGN §14.3 的本地 Ollama。"""
    for url in ("http://127.0.0.1:11434/v1", "http://[::1]:11434/v1",
                "http://localhost:11434/v1", "http://api.localhost/v1"):
        ai_assist._validate_endpoint({}, url)


def test_global_ip_literal_is_allowed():
    """公网 IP 字面量放行（不触 DNS 分支）。"""
    ai_assist._validate_endpoint({}, "http://93.184.216.34/v1")


def test_resolvable_public_domain_is_allowed(monkeypatch):
    """解析全为公网地址的域名放行。"""
    monkeypatch.setattr(ai_assist.socket, "getaddrinfo",
                        _resolved("api.example.invalid", "93.184.216.34"))
    ai_assist._validate_endpoint({}, "http://api.example.invalid/v1")


def test_unresolvable_hostname_is_rejected(monkeypatch):
    """getaddrinfo gaierror → 报「无法解析」拒收，不留到 urlopen 才炸。"""
    def gaierror(host_, *a, **k):
        raise socket.gaierror(11001, "getaddrinfo failed")
    monkeypatch.setattr(ai_assist.socket, "getaddrinfo", gaierror)
    with pytest.raises(ValueError, match="无法解析"):
        ai_assist._validate_endpoint({}, "http://no-such-host.example.invalid/v1")


def test_other_falsy_opt_in_values_still_reject_private():
    """opt-in 只认 is True：字符串 "1"、1、真值列表都放不进内网端点。"""
    for bad in ("1", 1, [True]):
        with pytest.raises(ValueError, match="allow_lan_endpoint"):
            ai_assist._validate_endpoint({"allow_lan_endpoint": bad},
                                         "http://10.0.0.2/v1")
