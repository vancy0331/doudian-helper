"""完全按录制代码填写每个字段"""
import os, re, glob
os.chdir(r"C:\Users\86153\Desktop\抖店上架助手")
from playwright.sync_api import sync_playwright

style_code = "DY2737"
from pathlib import Path
BASE_DIR = Path.cwd()
title = "2026夏季法式小香风两件套时尚套装女短袖上衣配半身裙套装"

with sync_playwright() as p:
    ctx = p.chromium.launch_persistent_context(
        user_data_dir="browser_data", headless=False,
        args=["--disable-blink-features=AutomationControlled"])
    page = ctx.new_page()
    page.goto("https://fxg.jinritemai.com/ffa/g/create", wait_until="domcontentloaded", timeout=15000)
    page.wait_for_timeout(8000)

    # Step1
    imgs = glob.glob(f"商品图/主图1x1/{style_code}*")
    page.locator("label").filter(has_text="商品正面图*上传主图").set_input_files(imgs[:5])
    page.wait_for_timeout(3000)
    page.locator("#pg-title-input").fill(title)
    page.wait_for_timeout(2000)
    page.locator("#pg-title-input").click(); page.wait_for_timeout(300)
    page.keyboard.press("Enter"); page.wait_for_timeout(3000)
    page.wait_for_selector(".style_predictContentWrapper__UdtJQ", timeout=10000)
    page.locator(".style_predictContentWrapper__UdtJQ svg").first.click()
    page.wait_for_timeout(2000)
    page.get_by_role("button", name="下一步").first.click(timeout=10000)
    page.wait_for_timeout(4000)

    for _ in range(3): page.keyboard.press("Escape"); page.wait_for_timeout(300)

    # 扫描连衣裙页面字段
    info = page.evaluate("""() => {
        let r = [];
        document.querySelectorAll('input, [class*="select"]').forEach(el => {
            let rect = el.getBoundingClientRect();
            if (rect.width < 10 || rect.height < 5 || el.type === 'hidden' || el.type === 'file') return;
            let label = '';
            let p = el.parentElement;
            for (let i=0; i<5 && p; i++) {
                let t = (p.innerText||'').trim().slice(0,50);
                if (t && t.length < 50) { label = t; break; }
                p = p.parentElement;
            }
            let val = el.value || el.innerText || '';
            r.push({y:Math.round(rect.y), label:label.slice(0,40), val:val.slice(0,25)});
        });
        return r.filter(x=>x.y>300).sort((a,b)=>a.y-b.y).slice(0,30);
    }""")
    print("=== 连衣裙页面字段 ===")
    for i in info:
        print(f"  y={i['y']} [{i['label'][:30]}] = {i['val'][:20]}")

    _sn = [0]
    def shot(label):
        page.screenshot(path=f"t4_{_sn[0]:02d}_{label}.png")
        print(f"t4_{_sn[0]:02d}_{label}.png")
        _sn[0] += 1

    shot("0进入详情页")

    # 短标题
    page.get_by_role("textbox", name="建议填写简明准确的标题内容，避免重复表达").fill("法式小香风两件套时尚套装")
    page.wait_for_timeout(300)
    shot("1短标题")

    # 品牌
    page.locator('//*[contains(text(),"品牌")]/ancestor::div[1]').first.click()
    page.wait_for_timeout(500)
    page.get_by_text("无品牌", exact=True).first.click(timeout=2000)
    page.wait_for_timeout(300)
    shot("2品牌")

    # 货号
    page.locator('//span[contains(text(),"货号")]/following::input[@placeholder="请输入"][1]').fill(style_code)
    page.wait_for_timeout(300)
    shot("3货号")

    # 上市时间（连衣裙可能没有这个字段）
    try:
        page.locator('//*[contains(text(),"上市时间")]/ancestor::div[3]//input[not(@id="pg-title-input")]').first.fill("2026")
        page.wait_for_timeout(300)
        shot("4上市时间")
    except:
        print("  上市时间字段不存在，跳过")

    # ====== 面料材质（按标签文字定位，不依赖固定ID） ======
    # 1. 点面料材质下拉
    page.locator('//*[contains(text(),"面料材质")]/following::div[contains(@class,"ecom-g-select-selector")][1]').click()
    page.wait_for_timeout(500)
    # 填搜索词 + force点正确选项（不用键盘选，防选错）
    page.locator('//*[contains(text(),"面料材质")]/following::input[@type="search"][1]').fill("聚酯纤维")
    page.wait_for_timeout(1500)
    try:
        page.get_by_text("聚酯纤维（涤纶）").first.click(force=True, timeout=3000)
    except:
        page.get_by_text("聚酯纤维", exact=True).first.click(force=True, timeout=3000)
    page.wait_for_timeout(300)
    page.wait_for_timeout(500)
    shot("5面料-聚酯纤维")

    # 2. 填85%
    page.locator('//*[contains(text(),"面料材质")]/following::input[@placeholder="请输入"][1]').fill("85")
    page.wait_for_timeout(300)
    shot("6面料-85%")

    # 3. 添加材质
    page.get_by_role("button", name="添加材质").click()
    page.wait_for_timeout(800)
    shot("7面料-添加材质")

    # 4. 第二个下拉选其他纤维
    page.locator('//*[contains(text(),"面料材质")]/following::div[contains(@class,"ecom-g-select-selector")]').last.click()
    page.wait_for_timeout(500)
    page.locator('//*[contains(text(),"面料材质")]/following::input[@type="search"][2]').fill("其他")
    page.wait_for_timeout(1500)
    page.keyboard.press("ArrowDown")
    page.wait_for_timeout(300)
    page.keyboard.press("Enter")
    page.wait_for_timeout(500)
    shot("8面料-其他纤维")

    # 5. 填15%
    page.locator('//*[contains(text(),"面料材质")]/following::input[@placeholder="请输入"]').nth(1).fill("15")
    page.wait_for_timeout(300)
    shot("9面料-15%")

    # 尺码表 + 从主图填入（按录制）
    sc = glob.glob("商品图/参考图/*尺码表*")
    if sc:
        # 删掉商详区已有图片（点垃圾桶图标）
        for _ in range(5):
            try:
                page.locator('[class*="delete"], [class*="trash"], [class*="Delete"], [class*="Trash"], [aria-label*="删除"], [aria-label*="delete"]').first.click(timeout=1000)
                page.wait_for_timeout(300)
            except: break
        page.locator('[id="goodsEditScrollContainer-图文信息"] label').filter(has_text="上传图片").set_input_files(sc[0])
        page.wait_for_timeout(2000)
    shot("10尺码表")
    page.get_by_role("button", name="从主图填入").click()
    page.wait_for_timeout(2000)
    shot("11从主图填入")

    # 3:4主图 - 暂时跳过

    # 价格库存tab
    page.get_by_text("价格库存").first.click()
    page.wait_for_timeout(2000)
    shot("12价格库存tab")

    # 发货模式 - radio按钮！
    page.get_by_role("radio", name="现货预售混合").check()
    page.wait_for_timeout(300)
    page.get_by_role("radio", name="小时").check()
    page.wait_for_timeout(300)
    page.get_by_role("checkbox", name="15天内", exact=True).check()
    page.wait_for_timeout(300)
    shot("13发货模式")

    # SKU创建（录制模式，动态SKU值）
    sku_list = ["上衣", "裙子", "套装"]  # TODO: 从表格读取
    page.locator(".anticon.anticon-down.ecom-g-cascader-picker-arrow > svg").first.click()
    page.wait_for_timeout(500)
    page.get_by_text("创建类型").click()
    page.wait_for_timeout(500)
    for i, sku in enumerate(sku_list):
        page.get_by_role("textbox", name="请输入规格值").fill(sku)
        page.wait_for_timeout(200)
        if i < len(sku_list) - 1:
            page.locator(".styles_addSKUNameEdit__u7cgM > span > .icon").first.click()
            page.wait_for_timeout(500)
            page.get_by_text("创建类型").click()
            page.wait_for_timeout(500)
        else:
            page.locator(".styles_addSKUNameEdit__u7cgM").click()
            page.wait_for_timeout(200)
            page.locator(".styles_addSKUNameEdit__u7cgM > span > .icon").first.click()
            page.wait_for_timeout(500)
    page.get_by_role("button", name="确定").click()
    page.wait_for_timeout(2000)
    # SKU图 - 每次取first（上传后DOM减少）
    imgs_sku = list((BASE_DIR / "商品图" / "主图1x1").glob(f"{style_code}*"))
    for i in range(len(sku_list)):
        f = str(imgs_sku[i])
        with page.expect_file_chooser() as fc:
            page.locator(".style_imgUpload__XG5rw .icon").first.click()
        fc.value.set_files(f)
        page.wait_for_timeout(1500)
    shot("14SKU+图")

    # 尺码信息tab
    page.get_by_text("尺码表").first.click()
    page.wait_for_timeout(1500)

    # 尺码大小（按最新录制）
    page.locator('[id="skuValue-尺码大小"] > .style_skuValueInput__fpVup > div > .ecom-g-cascader-picker > .anticon > svg').click()
    page.wait_for_timeout(500)
    for n in [7, 8, 9, 10]:
        page.locator(f"li:nth-child({n}) > .ecom-g-checkbox-wrapper > .ecom-g-checkbox > .ecom-g-checkbox-input").check()
        page.wait_for_timeout(200)
    page.get_by_role("button", name="确定 (4)").click()
    page.wait_for_timeout(1000)
    # 滚动到尺码区域
    page.evaluate("window.scrollTo(0, 3000)")
    page.wait_for_timeout(500)
    # 尺码大小区域内的备注框
    sizes_area = page.locator('[id="skuValue-尺码大小"]')
    sizes_area.get_by_role("textbox", name="备注").first.fill("80-95斤")
    page.wait_for_timeout(200)
    sizes_area.get_by_role("textbox", name="备注").nth(1).fill("95-105斤")
    page.wait_for_timeout(200)
    sizes_area.get_by_role("textbox", name="备注").nth(2).fill("105-115斤")
    page.wait_for_timeout(200)
    sizes_area.get_by_role("textbox", name="备注").nth(3).fill("115-125斤")
    page.wait_for_timeout(200)
    shot("16尺码大小")

    # 尺码模板
    page.locator("div").filter(has_text=re.compile(r"^一键复用尺码信息$")).nth(2).click()
    page.wait_for_timeout(500)
    try:
        page.get_by_text("128尺码").first.click(timeout=2000)
    except:
        page.get_by_text("4码").first.click(timeout=2000)
    page.wait_for_timeout(500)
    shot("17尺码模板")
    # 价格库存（按录制，支持任意数量SKU）
    prices = ["59.9", "69.9", "109.9"]  # TODO: 从表格读取
    stock = "5000"

    for i, sku_name in enumerate(sku_list):
        if i == 0:
            page.get_by_placeholder("价格").click()
            page.get_by_placeholder("价格").fill(prices[i])
            # 确认价格填对了
            val = page.get_by_placeholder("价格").input_value()
            if val != prices[i]:
                page.get_by_placeholder("价格").fill(prices[i])
            page.locator(".ecom-g-select.sp-select-all-tag > .ecom-g-select-selector").first.click()
            page.wait_for_timeout(300)
            page.locator(".select-all-checkbox > svg").click()
            page.wait_for_timeout(300)
            page.get_by_title(sku_name).click()
            page.wait_for_timeout(200)
            page.get_by_role("spinbutton", name="现货库存").fill("0")
            page.wait_for_timeout(200)
            page.get_by_role("spinbutton", name="预售库存").click()
            page.get_by_role("spinbutton", name="预售库存").fill(stock)
        else:
            page.locator(".ecom-g-select.ecom-g-select-multiple.ecom-g-select-show-arrow > .ecom-g-select-selector").first.click()
            page.wait_for_timeout(300)
            page.get_by_title(sku_name).click()
            page.locator(".anticon.anticon-check > svg").first.click()
            page.wait_for_timeout(200)
            page.get_by_placeholder("价格").click()
            page.get_by_placeholder("价格").fill(prices[i])
            val = page.get_by_placeholder("价格").input_value()
            if val != prices[i]:
                page.get_by_placeholder("价格").fill(prices[i])
        page.get_by_role("button", name="批量设置").click()
        page.wait_for_timeout(800)
    shot("15价格库存")


    # 下架+发布
    page.evaluate("window.scrollTo(0, 99999)"); page.wait_for_timeout(500)
    page.get_by_text("付款减库存").first.click(); page.wait_for_timeout(300)
    page.get_by_text("下架").first.click(); page.wait_for_timeout(500)
    shot("16下架")
    page.get_by_role("button", name="发布商品").first.click(timeout=5000)
    page.wait_for_timeout(3000)
    shot("17发布")

    ctx.close()
    print("\n请逐一查看截图")
