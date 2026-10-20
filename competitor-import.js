/* Kalodata XLSX import stays in memory; no upload, storage or background sync. */
(() => {
  const input = document.querySelector("#competitor-import-files");
  const clear = document.querySelector("#competitor-clear-import");
  const status = document.querySelector("#competitor-import-status");
  let original = null;
  const countryNames = { DE: "德国", FR: "法国", IT: "意大利" };

  function parseXml(text) {
    const doc = new DOMParser().parseFromString(text, "application/xml");
    if (doc.getElementsByTagName("parsererror").length) throw new Error("Excel XML 格式无效");
    return doc;
  }
  const nodes = (root, name) => [...root.getElementsByTagNameNS("*", name)];
  const readXml = async (zip, path) => {
    const file = zip.file(path);
    if (!file) throw new Error(`缺少必要工作表文件 ${path}`);
    return parseXml(await file.async("string"));
  };
  async function parseFile(file) {
    if (!/\.xlsx$/i.test(file.name) || file.size > 10 * 1024 * 1024) throw new Error("请选择不超过 10 MB 的 XLSX 文件");
    const match = file.name.match(/_(DE|FR|IT)(?:\b|[_ .-])/i);
    if (!match) throw new Error(`${file.name} 的文件名需包含 _DE、_FR 或 _IT，以识别国家`);
    const country = countryNames[match[1].toUpperCase()];
    const zip = await JSZip.loadAsync(await file.arrayBuffer());
    const book = await readXml(zip, "xl/workbook.xml");
    const rels = await readXml(zip, "xl/_rels/workbook.xml.rels");
    const sheet = nodes(book, "sheet").find((s) => s.getAttribute("name") === "LIST_PRODUCT");
    if (!sheet) throw new Error(`${file.name} 缺少 LIST_PRODUCT 工作表`);
    const rid = sheet.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id");
    const rel = nodes(rels, "Relationship").find((r) => r.getAttribute("Id") === rid);
    if (!rel) throw new Error("无法定位商品工作表");
    let target = rel.getAttribute("Target").replace(/^\//, "");
    if (!target.startsWith("xl/")) target = `xl/${target}`;
    if (!/^xl\/worksheets\/[\w.-]+\.xml$/.test(target)) throw new Error("工作表路径不受支持");
    const sharedDoc = zip.file("xl/sharedStrings.xml") ? await readXml(zip, "xl/sharedStrings.xml") : null;
    const shared = sharedDoc ? nodes(sharedDoc, "si").map((s) => nodes(s, "t").map((t) => t.textContent).join("")) : [];
    const doc = await readXml(zip, target);
    const rows = nodes(doc, "row").map((row) => {
      const result = [];
      for (const cell of nodes(row, "c")) {
        const letters = (cell.getAttribute("r") || "").match(/^[A-Z]+/)?.[0];
        if (!letters) continue;
        const index = [...letters].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
        const kind = cell.getAttribute("t");
        const v = nodes(cell, "v")[0]?.textContent;
        result[index] = kind === "inlineStr" ? nodes(cell, "t").map((t) => t.textContent).join("") : kind === "s" ? shared[Number(v)] : v ?? "";
      }
      return result;
    });
    const headers = rows[0] || [];
    const required = ["日期范围", "商品名称", "类目", "销量", "平均销售价(¥)", "成交金额(¥)"];
    if (required.some((name) => !headers.includes(name))) throw new Error(`${file.name} 缺少必要的 Kalodata 字段`);
    const field = (row, name) => row[headers.indexOf(name)];
    const number = (v) => v === "" || v == null || v === "-" || !Number.isFinite(Number(v)) ? null : Number(v);
    const records = rows.slice(1).filter((row) => field(row, "商品名称")).map((row, i) => ({
      brand: String(field(row, "商品名称")), category: String(field(row, "类目") || "--"), model: "--", country, platform: "TikTok Shop",
      sales: number(field(row, "销量")), averageSalePrice: number(field(row, "平均销售价(¥)")), revenue: number(field(row, "成交金额(¥)")),
      price: number(field(row, "价格(¥)")), listedAt: field(row, "上架时间"), commission: field(row, "佣金比例"),
      revenueGrowth: field(row, "成交金额增长率"), liveRevenue: number(field(row, "直播成交金额(¥)")),
      videoRevenue: number(field(row, "视频成交金额(¥)")), cardRevenue: field(row, "商品卡成交金额"),
      creators: number(field(row, "达人数量")), creatorOrderRate: field(row, "达人出单率"),
      imageUrl: field(row, "图片链接"), kalodataUrl: field(row, "Kalodata详情页链接"), tiktokUrl: field(row, "TikTok链接"),
      period: String(field(row, "日期范围") || ""), rank: i + 1, change: null, rankDerived: false,
    }));
    if (!records.length) throw new Error(`${file.name} 没有商品记录`);
    return records;
  }

  input.addEventListener("change", async () => {
    const files = [...input.files];
    if (!files.length) return;
    input.disabled = true;
    status.textContent = "正在本地读取文件……";
    try {
      if (!state.data) throw new Error("请等待演示数据加载完成后重试");
      const groups = await Promise.all(files.map(parseFile));
      const records = groups.flat();
      if (!original) original = state.data.competitors;
      state.data.competitors = records;
      state.data.meta.competitorLocalImport = true;
      for (const id of ["competitor-search", "competitor-country", "competitor-platform"]) document.querySelector(`#${id}`).value = "";
      document.querySelector('#competitor-sort option[value="rank"]').textContent = "导出顺序从前到后";
      renderCompetitors();
      clear.hidden = false;
      const periods = [...new Set(records.map((r) => r.period).filter(Boolean))];
      status.textContent = `已在本机临时导入 ${files.length} 个文件，共 ${records.length} 条；统计日期：${periods.join("、") || "未知"}。刷新后需重新导入。`;
    } catch (error) {
      status.textContent = `导入失败：${error.message}。未替换现有数据。`;
    } finally { input.disabled = false; input.value = ""; }
  });
  clear.addEventListener("click", () => {
    if (!original || !state.data) return;
    state.data.competitors = original;
    delete state.data.meta.competitorLocalImport;
    original = null;
    clear.hidden = true;
    document.querySelector('#competitor-sort option[value="rank"]').textContent = "排名从前到后";
    status.textContent = "已恢复虚构演示数据。";
    renderCompetitors();
  });
  document.querySelector("#refresh-button").addEventListener("click", () => {
    original = null;
    clear.hidden = true;
    document.querySelector('#competitor-sort option[value="rank"]').textContent = "排名从前到后";
    status.textContent = "刷新后将恢复演示数据，请重新导入文件。";
  });
})();
