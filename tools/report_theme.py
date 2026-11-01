"""三份对外报告（投稿体检 / 审计 / 规格页）共用的视觉基础。

配色与字体与 Web 端 templates/base.html 保持一致：
同一款产品不应该在浏览器里出现两套视觉语言，尤其在报告被嵌进结果页 iframe 时。
"""

FONT = 'Inter,"PingFang SC","Microsoft YaHei",system-ui,-apple-system,sans-serif'

BASE = """
:root{--bg:#f4f6f3;--surface:#fff;--soft:#f8faf7;--ink:#18332e;--ink2:#52655f;
--ink3:#64756e;--line:#dfe6e1;--line2:#edf1ed;--accent:#176b57;--dark:#0e4f41;
--accent-soft:#e6f2ed;--gold:#8a6626;--gold-bg:#f7f0e3;--danger:#a44b3f;--danger-bg:#fbede9}
*{box-sizing:border-box}
body{margin:0;padding:32px 24px 60px;background:var(--bg);color:var(--ink);
font-family:FONT_PLACEHOLDER;line-height:1.7;-webkit-font-smoothing:antialiased}
.wrap{max-width:920px;margin:0 auto}
a{color:var(--accent)}
header{margin-bottom:26px;padding-bottom:20px;border-bottom:1px solid var(--line)}
h1{font-family:"Songti SC","Noto Serif SC",SimSun,serif;font-size:25px;font-weight:600;
margin:0 0 8px;letter-spacing:.6px;line-height:1.4}
.sub{color:var(--ink2);font-size:13.5px;margin:0 0 7px}
.meta{color:var(--ink3);font-size:12.5px;margin:0}
section{margin:0 0 18px;background:var(--surface);border:1px solid var(--line);
border-radius:14px;padding:20px 22px}
section h2{font-size:14px;font-weight:600;margin:0 0 14px;padding-left:10px;
border-left:3px solid var(--accent);line-height:1.4}
table{width:100%;border-collapse:collapse;font-size:13.5px}
th{text-align:left;font-weight:500;color:var(--ink3);font-size:12px;letter-spacing:.3px;
padding:0 10px 9px 0;border-bottom:1px solid var(--line2)}
td{padding:9px 10px 9px 0;border-top:1px solid var(--line2);vertical-align:top}
tr:first-child>td{border-top:none}
footer{margin-top:34px;color:var(--ink3);font-size:12px;text-align:center;line-height:1.9}
.tag{display:inline-block;font-size:11.5px;padding:2px 9px;border-radius:999px;
white-space:nowrap;background:var(--line2);color:var(--ink2)}
.good{background:var(--accent-soft);color:var(--accent)}
.bad{background:var(--danger-bg);color:var(--danger)}
.warn{background:var(--danger-bg);border-color:#efcfc7}
.warn h2{border-left-color:var(--danger)}
@media print{
  body{padding:0;background:#fff;font-size:11pt}
  section{border:none;border-radius:0;padding:0 0 10px;break-inside:avoid}
  header{margin-bottom:14px}
  .no-print{display:none!important}
  a{color:inherit;text-decoration:none}
}
""".replace("FONT_PLACEHOLDER", FONT)
