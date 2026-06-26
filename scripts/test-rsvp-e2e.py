#!/usr/bin/env python3
"""E2E test: fill and submit RSVP form on live site via Playwright."""

import time
from playwright.sync_api import sync_playwright

URL = "https://our-wedding-day.github.io/info/#rsvp"


def dismiss_music_gate(page):
    skip = page.locator("#musicGateSkip")
    try:
        if skip.is_visible(timeout=3000):
            skip.click()
            page.wait_for_timeout(400)
    except Exception:
        pass


def fill_attending_flow(page, contact, comment):
    page.goto(URL, wait_until="networkidle", timeout=60000)
    dismiss_music_gate(page)
    page.fill("#name", "Тест Cursor Автотест")
    page.fill("#contact", contact)
    page.check('input[name="attend"][value="Так, буду"]')
    page.click("#rsvpNext")
    page.wait_for_timeout(400)

    page.select_option("#guests", "2")
    page.check('input[name="welcome"][value="Так"]')
    page.check('input[name="church"][value="Так"]')
    page.check('input[name="room"][value="Так"]')
    page.fill("#travelNotes", "Окремі ліжка (автотест)")
    page.fill("#foodNotes", "Без горіхів (автотест)")
    page.wait_for_timeout(400)

    page.fill("#comment", comment)


def fill_decline_flow(page, contact):
    page.goto(URL, wait_until="networkidle", timeout=60000)
    dismiss_music_gate(page)
    page.fill("#name", "Тест Cursor Відмова")
    page.fill("#contact", contact)
    page.check('input[name="attend"][value="На жаль, ні"]')
    page.wait_for_selector("#rsvpDeclineHint.is-visible")
    page.click("#rsvpNext")
    page.wait_for_timeout(400)
    page.fill("#comment", "Не зможу — автотест Playwright")


def submit_and_collect(page, results, label):
    responses = []

    def on_response(r):
        if "script.google.com" in r.url or "googleusercontent.com" in r.url:
            responses.append(f"{r.status} {r.url[:80]}")

    page.on("response", on_response)
    page.click("#submitBtn")

    try:
        page.wait_for_function(
            """() => {
                const s = document.getElementById('formSuccess');
                const e = document.getElementById('formError');
                return (s && s.style.display === 'block') || (e && e.style.display === 'block');
            }""",
            timeout=30000,
        )
    except Exception as exc:
        results.append((f"{label}: wait", f"TIMEOUT: {exc}"))

    for r in responses:
        results.append((f"{label}: network", r))

    success = page.locator("#formSuccess")
    error = page.locator("#formError")
    if success.is_visible():
        title = page.locator("#formSuccessTitle").inner_text()
        text = page.locator("#formSuccessText").inner_text()
        results.append((f"{label}: UI", f"SUCCESS — {title}"))
        results.append((f"{label}: text", text[:160]))
    elif error.is_visible():
        results.append((f"{label}: UI", f"ERROR — {error.inner_text()}"))
    else:
        results.append((f"{label}: UI", "UNKNOWN"))


def run():
    results = []
    ts = int(time.time())
    email_create = f"rsvp-test-{ts}@example.com"
    phone_create = f"+38067{ts % 10000000:07d}"
    email_decline = f"rsvp-decline-{ts}@example.com"

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)

        page = browser.new_page()
        fill_attending_flow(page, email_create, "Автотест Playwright — можна видалити")
        submit_and_collect(page, results, "Create email")

        page = browser.new_page()
        fill_attending_flow(page, email_create, "Оновлено автотестом Playwright")
        submit_and_collect(page, results, "Update email")

        page = browser.new_page()
        fill_attending_flow(page, phone_create, "Автотест телефон — можна видалити")
        submit_and_collect(page, results, "Create phone")

        page = browser.new_page()
        fill_decline_flow(page, email_decline)
        submit_and_collect(page, results, "Decline")

        browser.close()

    return results


if __name__ == "__main__":
    for label, value in run():
        print(f"{label}: {value}")
