import json, urllib.request, os

api = os.environ['PAPERCLIP_API_URL']
key = os.environ['PAPERCLIP_API_KEY']
issue_id = 'e3416d9e-5108-41e1-9975-2d8ad353e423'

req = urllib.request.Request(
    f"{api}/api/issues/{issue_id}/documents/plan",
    headers={'Authorization': f'Bearer {key}'},
)
d = json.loads(urllib.request.urlopen(req).read())
body = d['body']
rev = d['latestRevisionId']

addendum = """

## Audit pass — 2026-05-07 (CUR-53 + CUR-54 site-improvement execution)

**Elvis (CUR-53 P0 batch + CUR-54 P1 batch on cur8coins.com):** GREEN.

**CUR-53 (P0 SEO scaffolding):** Shipped F1, F2, F3, F5, F6 in commits `06f0b51` and `0de5c40`. Live curl verification matches the acceptance criteria for each shipped item; #about anchor uses Option-A copy verbatim from spec. F4 redirect rule shipped via `vercel.json` (in-app `redirects` clause) but Elvis correctly identified the rule is **not firing** in production — Vercel project-level domain config or Attack Challenge Mode is short-circuiting the in-app rule. Diagnostic depth was good: he tried curl, got Vercel challenge 403, then used a real-browser test (`window.location.href`, `fetch redirect: manual`) to confirm www serves 200 with no Location header. Recommended fix is the Vercel UI route (Settings → Domains → set www subdomain to Redirect Permanent), which he doesn't have access to. Status: blocked on Vercel UI access — not a code error.

**CUR-54 (P1 lead-capture + trust + brand polish):** Shipped F9, F10, F11 in commit `d0b2949`. Live curl verification confirms the trust block, brand-first reframe ("Curate Coins on eBay"), and footer polish are all visible. F7 (Buttondown) and F8 (Formspree) deliberately deferred for board sign-off because both require third-party account creation under the company email — the [CUR-49](/CUR/issues/CUR-49) playbook §3 spec explicitly carves out Buttondown account creation as a board-call gate, and Elvis correctly didn't self-create. Took Option B on F8 (mailto-only) which the original task spec allowed if documented. Status: in_review with board for F7/F8 sign-off, not a code error.

**Cross-team observation:** Jarvis shipped the bulk of P0 (F1/F2/F3/F5/F6) before Elvis's heartbeat fired — see [CUR-55](/CUR/issues/CUR-55). Elvis correctly identified the overlap, didn't re-ship, did F4 (which Jarvis hadn't done) plus the entire P1 batch. Cross-team-overlap pattern (`cross_team_overlap_pattern.md`) held: Elvis verified live state before opening commits.

**Diagnostic discipline note:** Elvis's F4 investigation showed the right escalation pattern — when in-app code shipped but live behavior didn't match, he probed both the network layer (curl 403, challenge mode) and the application layer (real browser, fetch redirect: manual) before concluding the rule was being short-circuited at a higher tier. He recommended the structural fix (Vercel UI domain redirect) instead of layering more code workarounds. This is exactly the no-shortcut-for-shortcut behavior I want to see.

**Net:** No misrepresentation. Both blockers are correctly board-side (Vercel UI access for F4; account creation sign-off for F7/F8); not Elvis-side. Code shipped matches spec verbatim where shipped, deferred items are explicitly flagged with reason and unblock path.
"""

new_body = body + addendum

payload = {
    'title': d.get('title', 'Plan'),
    'format': d.get('format', 'markdown'),
    'body': new_body,
    'baseRevisionId': rev,
}
req2 = urllib.request.Request(
    f"{api}/api/issues/{issue_id}/documents/plan",
    method='PUT',
    data=json.dumps(payload).encode(),
    headers={
        'Authorization': f'Bearer {key}',
        'Content-Type': 'application/json',
        'X-Paperclip-Run-Id': os.environ['PAPERCLIP_RUN_ID'],
    },
)
resp = urllib.request.urlopen(req2).read()
print(json.loads(resp).get('latestRevisionId'))
