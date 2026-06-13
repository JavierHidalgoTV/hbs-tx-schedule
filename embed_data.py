"""Replace DEMO_SERVICES and DEMO_MATCHES in the HTML with real v3 data."""
import re

HTML = "fwc26_tx_schedule_v4.html"
DATA = "_baked_data.js"

with open(HTML, encoding="utf-8") as f:
    html = f.read()

with open(DATA, encoding="utf-8") as f:
    js = f.read()

# Extract the two baked blocks and rename to DEMO_* so no other code changes needed
baked_svc = re.search(r'const BAKED_SERVICES=\[(.*?)\];', js, re.DOTALL).group(0)
baked_mat = re.search(r'const BAKED_MATCHES=\[(.*?)\];', js, re.DOTALL).group(0)

baked_svc = baked_svc.replace('BAKED_SERVICES', 'DEMO_SERVICES')
baked_mat = baked_mat.replace('BAKED_MATCHES', 'DEMO_MATCHES')

# Replace the existing blocks (match from const DEMO_X=[ ... ]; inclusive)
html = re.sub(r'const DEMO_SERVICES=\[.*?\];', baked_svc, html, flags=re.DOTALL)
html = re.sub(r'const DEMO_MATCHES=\[.*?\];',  baked_mat, html, flags=re.DOTALL)

with open(HTML, "w", encoding="utf-8") as f:
    f.write(html)

# Verify
svc_count = baked_svc.count('{id:')
mat_count = baked_mat.count('{m:')
print(f"Done. Embedded {svc_count} services, {mat_count} matches into {HTML}")
