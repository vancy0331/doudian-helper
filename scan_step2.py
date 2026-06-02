"""扫描详情页：每个label和对应的input"""
import os, re, glob
os.chdir(r"C:\Users\86153\Desktop\抖店上架助手")
from playwright.sync_api import sync_playwright

with sync_playwright() as p:
    ctx = p.chromium.launch_persistent_context(
        user_data_dir="browser_data", headless=False,
        args=["--disable-blink-features=AutomationControlled"])
    page = ctx.new_page()
    page.goto("https://fxg.jinritemai.com/ffa/g/create", wait_until="domcontentloaded", timeout=15000)
    page.wait_for_timeout(8000)

    imgs = glob.glob("商品图/主图1x1/DY2737*")
    page.locator("label").filter(has_text="商品正面图*上传主图").set_input_files(imgs[:5])
    page.wait_for_timeout(3000)
    page.locator("#pg-title-input").fill("2026夏季法式小香风两件套时尚套装女短袖上衣配半身裙套装")
    page.wait_for_timeout(2000)
    page.locator("#pg-title-input").click(); page.wait_for_timeout(300)
    page.keyboard.press("Enter"); page.wait_for_timeout(3000)
    page.wait_for_selector(".style_predictContentWrapper__UdtJQ", timeout=10000)
    page.locator(".style_predictContentWrapper__UdtJQ svg").first.click()
    page.wait_for_timeout(2000)
    page.get_by_role("button", name="下一步").first.click(timeout=10000)
    page.wait_for_timeout(4000)
    print("进入详情页\n")

    # 保存完整HTML
    with open("scan_step2.html", "w", encoding="utf-8") as f:
        f.write(page.content())
    print("已保存 scan_step2.html")

    # 提取所有label-input配对
    info = page.evaluate("""() => {
        let result = [];
        let all = document.querySelectorAll('input, textarea, [class*="select"]');
        for (let e of all) {
            let r = e.getBoundingClientRect();
            if (r.width < 10 || r.height < 5) continue;
            if (e.type === 'hidden' || e.type === 'file') continue;
            // 找最近的前面标签文字
            let label = '';
            let prev = e.previousElementSibling;
            for (let i=0; i<5 && prev; i++) {
                let t = (prev.innerText||'').trim().slice(0,30);
                if (t) { label = t; break; }
                prev = prev.previousElementSibling;
            }
            // 找父容器的文本
            if (!label) {
                let p = e.parentElement;
                for (let i=0; i<4 && p; i++) {
                    let t = (p.innerText||'').trim().slice(0,40);
                    if (t && t.length < 40) { label = t; break; }
                    p = p.parentElement;
                }
            }
            result.push({
                tag: e.tagName,
                type: e.type || '',
                placeholder: (e.placeholder||'').slice(0,40),
                label: label.slice(0,40),
                cls: (e.className?.toString?.()||'').slice(0,40),
                id: (e.id||'').slice(0,30),
                y: Math.round(r.y)
            });
        }
        return result.filter(r => r.y > 300).sort((a,b)=>a.y-b.y).slice(0,40);
    }""")

    for i in info:
        print(f"  y={i['y']} [{i['tag']}] id={i['id']} ph='{i['placeholder']}' label='{i['label']}'")

    ctx.close()
    print("\n把这些输出发给我")
