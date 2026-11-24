"""Clicks through the demo in Chromium (pip install playwright && playwright install chromium). Saves screenshots/."""
import asyncio, os, sys
from playwright.async_api import async_playwright
HERE = os.path.dirname(os.path.abspath(__file__))
APP = 'file://' + os.path.join(HERE, '..', 'index.html')
OUT = os.path.join(HERE, '..', 'screenshots'); os.makedirs(OUT, exist_ok=True)
CHROME = os.environ.get('CHROME')  # optional path to a Chrome binary

async def main():
    async with async_playwright() as p:
        b = await p.chromium.launch(executable_path=CHROME, args=['--no-sandbox']) if CHROME else await p.chromium.launch(args=['--no-sandbox'])
        ctx = await b.new_context(viewport={'width': 1440, 'height': 900})
        page = await ctx.new_page()
        errors = []
        page.on('console', lambda m: errors.append(m.text) if m.type == 'error' else None)
        page.on('pageerror', lambda e: errors.append('PAGEERROR ' + str(e)))
        await page.goto(APP)
        await page.wait_for_selector('table.data tbody tr')
        # --- stock: two column groups ---
        await page.click('[data-fam="NPC-X"]')
        groups = await page.locator('thead tr.groups th').all_inner_texts()
        print('groups:', groups); assert [g.strip().lower() for g in groups] == ['stock and booking', 'usage']
        usage_cols = await page.locator('thead tr.cols th.usage').all_inner_texts()
        print('usage columns:', [u.split('\n')[0].strip(' ▾') for u in usage_cols])
        await page.screenshot(path=f'{OUT}/1_stock.png')
        # filter Status = in stock
        await page.click('.th-btn[data-col="c1"]'); await page.wait_for_selector('.filter-pop')
        await page.uncheck('.filter-pop [data-all]')
        labels = page.locator('.filter-pop .vals label')
        for i in range(await labels.count()):
            if (await labels.nth(i).inner_text()).startswith('in stock'): await labels.nth(i).locator('input').check()
        await page.click('.filter-pop [data-ok]')
        st = await page.locator('table.data tbody tr td:nth-child(2)').all_inner_texts()
        assert st and all(x.strip() == 'in stock' for x in st); n_in = len(st)
        # book from drawer
        await page.locator('table.data tbody tr').first.click(); await page.wait_for_selector('.drawer')
        await page.screenshot(path=f'{OUT}/2_drawer.png')
        await page.click('[data-act="book"]'); await page.wait_for_selector('.modal [data-proj]')
        await page.select_option('.modal [data-proj]', 'Training'); await page.click('.modal .btn.primary')
        await page.wait_for_selector('.toast'); print('toast:', await page.locator('.toast').first.inner_text())
        assert await page.locator('table.data tbody tr').count() == n_in - 1
        # --- overview: compact cards + expand ---
        await page.click('.topbar [data-tab="overview"]'); await page.wait_for_selector('#chBookings')
        assert await page.locator('.scard .details').count() == 0, 'cards compact by default'
        await page.locator('[data-more="NPC-X"]').click(); await page.wait_for_selector('.scard.open .details')
        print('expanded card:', (await page.locator('.scard.open .details').inner_text()).replace('\n', ' | '))
        # multi-select types + period
        await page.click('[data-pick="types"] button[data-v="NPC-X"]'); await page.click('[data-pick="types"] button[data-v="NPC-W"]')
        await page.click('[data-pick="period"] button[data-v="3"]'); await page.wait_for_selector('#chBookings')
        sel = await page.locator('[data-pick="types"] button[aria-pressed="true"]').all_inner_texts(); print('types selected:', sel); assert sel == ['NPC-X', 'NPC-W']
        h = await page.locator('.pair .card-h').first.inner_text(); print('chart header:', h.replace('\n', ' | ')); assert 'week' in h.lower()
        await page.click('[data-pick="period"] button[data-v="all"]'); await page.wait_for_selector('#chStock')
        await page.wait_for_timeout(300); await page.screenshot(path=f'{OUT}/3_overview.png', full_page=True)
        await page.click('[data-pick="types"] button[data-v=""]')
        # --- phone ---
        await page.click('.topbar [data-tab="phone"]'); await page.wait_for_selector('.phone-frame .pitem')
        mine_status = await page.locator('.phone-frame .phone > .phone-body .pitem .chip[class*="st-"]').count()
        print('my probes:', await page.locator('.phone-frame .phone > .phone-body .pitem').count(), 'status chips in list:', mine_status); assert mine_status == 0
        print('groups:', await page.locator('.phone-frame .phone > .phone-body .pbox-h span:first-child').all_inner_texts())
        await page.screenshot(path=f'{OUT}/4_phone_mine.png')
        await page.locator('.phone-frame .phone > .phone-body .pitem').first.click(); await page.wait_for_selector('.sheet [data-save]')
        await page.locator('.sheet .seg[data-role]').first.locator('button[data-v="good"]').click()
        await page.fill('.sheet [data-note]', 'checked in SEM, tip fine'); await page.click('.sheet [data-save]')
        await page.wait_for_selector('.toast'); hist = await page.locator('.sheet .history').inner_text()
        print('history:', hist.split('\n')[0]); assert 'checked in SEM, tip fine' in hist
        await page.click('.sheet [data-back]')
        await page.click('.phone-nav [data-ptab="book"]'); await page.select_option('[data-bookfam]', 'NPC-X')
        item = page.locator('.phone > .phone-body .pitem').first
        txt = (await item.inner_text()).replace('\n', ' | '); print('booking list item:', txt); assert 'El.' in txt and 'f ' in txt
        await page.screenshot(path=f'{OUT}/5_phone_book.png')
        await item.click(); await page.wait_for_selector('[data-confirm-book]')
        await page.locator('.sheet .seg[data-proj] button').nth(1).click(); await page.click('[data-confirm-book]')
        await page.wait_for_selector('.toast'); print('phone booking:', await page.locator('.toast').last.inner_text())
        # --- data checks: demo without new columns, then setup as operations ---
        await page.click('.topbar [data-tab="checks"]'); await page.wait_for_selector('.check')
        await page.click('#demoLegacy'); await page.wait_for_selector('#setup')
        await page.select_option('#persona', '1'); await page.wait_for_selector('#setup:not([disabled])')
        await page.click('#setup'); await page.wait_for_selector('.modal .plan')
        await page.screenshot(path=f'{OUT}/6_setup_preview.png')
        await page.click('.modal .btn.primary'); await page.wait_for_selector('text=All new columns and the Log are in place')
        print('setup:', await page.locator('.toast').last.inner_text())
        await page.screenshot(path=f'{OUT}/7_checks.png', full_page=True)
        # --- real phone width ---
        m = await b.new_context(viewport={'width': 390, 'height': 844}, device_scale_factor=2, is_mobile=True, has_touch=True)
        mp = await m.new_page(); mp.on('pageerror', lambda e: errors.append('MOBILE ' + str(e)))
        await mp.goto(APP); await mp.wait_for_selector('.phone.full .pitem'); await mp.screenshot(path=f'{OUT}/8_mobile.png')
        # --- dark OS setting still shows the light design ---
        d = await b.new_context(viewport={'width': 1440, 'height': 900}, color_scheme='dark')
        dp = await d.new_page(); await dp.goto(APP + '#overview'); await dp.wait_for_selector('#chBookings')
        bg = await dp.evaluate("getComputedStyle(document.body).backgroundColor"); print('body background with dark OS setting:', bg); assert bg == 'rgb(246, 247, 249)'
        print('ERRORS:', errors); assert not errors
        await b.close()
asyncio.run(main())
