"""扫描面料材质附近的DOM结构"""
import os; os.chdir(r"C:\Users\86153\Desktop\抖店上架助手")
from playwright.sync_api import sync_playwright

with sync_playwright() as p:
    ctx = p.chromium.launch_persistent_context(
        user_data_dir="browser_data", headless=False,
        args=["--disable-blink-features=AutomationControlled"])
    page = ctx.new_page()
    page.goto("https://fxg.jinritemai.com/ffa/g/create", wait_until="domcontentloaded", timeout=15000)
    page.wait_for_timeout(8000)

    # Step1 - 快速过到详情页
    import glob
    imgs = glob.glob("商品图/主图1x1/DY2737*")
    page.locator("label").filter(has_text="商品正面图*上传主图").set_input_files(imgs[:5])
    page.wait_for_timeout(3000)
    page.locator("#pg-title-input").fill("2026夏季法式小香风两件套时尚套装")
    page.wait_for_timeout(2000)
    page.locator("#pg-title-input").click(); page.wait_for_timeout(300)
    page.keyboard.press("Enter"); page.wait_for_timeout(3000)
    try:
        page.wait_for_selector(".style_predictContentWrapper__UdtJQ", timeout=5000)
        page.locator(".style_predictContentWrapper__UdtJQ svg").first.click()
        page.wait_for_timeout(2000)
    except: pass
    page.get_by_role("button", name="下一步").first.click(timeout=10000)
    page.wait_for_timeout(4000)

    print("进入详情页\n")

    # 扫描面料材质附近所有元素
    result = page.evaluate("""() => {
        // 找到"面料材质"文字所在元素
        let all = document.querySelectorAll('*');
        let matEl = null;
        for (let e of all) {
            if (e.childNodes.length === 1 && e.childNodes[0].nodeType === 3 && e.innerText?.trim() === '面料材质') {
                matEl = e; break;
            }
        }
        if (!matEl) {
            // 模糊匹配
            for (let e of all) {
                let t = e.innerText?.trim();
                if (t === '面料材质') { matEl = e; break; }
            }
        }
        if (!matEl) return 'NOT FOUND';

        let r = matEl.getBoundingClientRect();
        console.log('面料材质位置:', r.x, r.y);

        // 找面料材质附近50-600px范围内的所有input/select/div
        let nearby = [];
        for (let e of all) {
            let er = e.getBoundingClientRect();
            if (er.y < r.y - 50 || er.y > r.y + 600) continue;
            if (er.x < r.x - 100 || er.x > r.x + 900) continue;
            if (er.width < 10 || er.height < 5) continue;

            let tag = e.tagName;
            let cls = (e.className?.toString?.() || '').slice(0,60);
            let id = (e.id || '').slice(0,30);
            let txt = (e.innerText || '').trim().slice(0,40);
            let ph = (e.placeholder || '').slice(0,40);
            let tp = (e.type || '').slice(0,20);
            let role = (e.getAttribute('role') || '').slice(0,20);

            nearby.push({
                tag, id, type: tp, role,
                cls: cls.slice(0,50),
                text: txt.slice(0,30),
                ph: ph.slice(0,30),
                x: Math.round(er.x), y: Math.round(er.y),
                w: Math.round(er.width), h: Math.round(er.height)
            });
        }
        return nearby.sort((a,b) => a.y - b.y).slice(0,40);
    }""")

    print("面料材质附近元素（从上到下）：")
    for el in result:
        tag_info = f"<{el['tag']}>"
        if el['id']: tag_info += f" id={el['id']}"
        if el['type']: tag_info += f" type={el['type']}"
        if el['role']: tag_info += f" role={el['role']}"
        print(f"  y={el['y']} x={el['x']} {tag_info}")
        if el['cls']: print(f"    class={el['cls']}")
        if el['text']: print(f"    text='{el['text']}'")
        if el['ph']: print(f"    placeholder='{el['ph']}'")

    ctx.close()
