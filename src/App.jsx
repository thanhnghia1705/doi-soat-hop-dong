import React, { useMemo, useState } from "react";
import * as XLSX from "xlsx";
import { Upload, FileSpreadsheet, CheckCircle2, AlertCircle, Download, RefreshCw, Database, ShieldCheck } from "lucide-react";

const REQUIRED_FIELDS = {
  orders: [
    { key: "orderDate", label: "Ngày đơn hàng" },
    { key: "pharmacyCode", label: "Mã nhà thuốc" },
    { key: "pharmacyName", label: "Tên nhà thuốc" },
    { key: "productCode", label: "Mã sản phẩm" },
    { key: "productName", label: "Tên sản phẩm" },
    { key: "sales", label: "Doanh số" },
  ],
  products: [
    { key: "productCode", label: "Mã sản phẩm" },
    { key: "productName", label: "Tên sản phẩm" },
  ],
  contracts: [
    { key: "contractDate", label: "Ngày đăng ký hợp đồng" },
    { key: "pharmacyCode", label: "Mã nhà thuốc" },
    { key: "pharmacyName", label: "Tên nhà thuốc" },
    { key: "commitment", label: "Mức doanh số cam kết" },
  ],
};

const FIELD_LABELS = {
  orderDate: "Ngày đơn hàng",
  pharmacyCode: "Mã nhà thuốc",
  pharmacyName: "Tên nhà thuốc",
  productCode: "Mã sản phẩm",
  productName: "Tên sản phẩm",
  sales: "Doanh số",
  contractDate: "Ngày đăng ký hợp đồng",
  commitment: "Mức doanh số cam kết",
};

const SYNONYMS = {
  orderDate: [
    "ngaydonhang",
    "ngaydon",
    "ngayban",
    "ngayhoadon",
    "documentdate",
    "orderdate",
  ],
  pharmacyCode: [
    "manhathuoc",
    "ma nhathuoc",
    "macuahang",
    "makhachhang",
    "customercode",
    "custcode",
    "pharmacycode",
    "ma",
  ],
  pharmacyName: [
    "tennhathuoc",
    "ten nhathuoc",
    "tencuahang",
    "tenkhachhang",
    "customername",
    "pharmacyname",
    "ten",
  ],
  productCode: [
    "masanpham",
    "ma sanpham",
    "sku",
    "productcode",
    "itemcode",
    "msp",
  ],
  productName: [
    "tensanpham",
    "ten sanpham",
    "productname",
    "itemname",
    "tenhang",
  ],
  sales: [
    "doanhso",
    "doanh thu",
    "thanhtien",
    "giatribanhang",
    "sales",
    "amount",
    "revenue",
    "netsales",
  ],
  contractDate: [
    "ngaydangkyhopdong",
    "ngaykyhopdong",
    "ngaydangky",
    "contractdate",
    "registrationdate",
    "ngayhieuluc",
  ],
  commitment: [
    "muccamket",
    "muccamkethopdong",
    "doanhsocamket",
    "contractcommitment",
    "commitment",
    "targetsales",
    "camket",
  ],
};

const SHEET_PREVIEW_LIMIT = 12;

function normalizeText(value) {
  return String(value ?? "")
    .trim()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/đ/g, "d")
    .replace(/Đ/g, "D")
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
}

function autoMapHeaders(headers, requiredFields) {
  const mapped = {};

  requiredFields.forEach((field) => {
    const candidates = SYNONYMS[field.key] || [];
    const found = headers.find((header) => {
      const normalized = normalizeText(header);
      return candidates.some((candidate) => normalized === normalizeText(candidate) || normalized.includes(normalizeText(candidate)));
    });
    mapped[field.key] = found || "";
  });

  return mapped;
}

function excelDateToJSDate(serial) {
  const parsed = XLSX.SSF.parse_date_code(serial);
  if (!parsed) return null;
  return new Date(parsed.y, parsed.m - 1, parsed.d);
}

function parseFlexibleDate(value) {
  if (value === null || value === undefined || value === "") return null;

  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }

  if (typeof value === "number") {
    const d = excelDateToJSDate(value);
    return d && !Number.isNaN(d.getTime()) ? d : null;
  }

  const raw = String(value).trim();
  if (!raw) return null;

  const clean = raw.replace(/\./g, "/").replace(/-/g, "/");
  const parts = clean.split("/").map((x) => x.trim());

  if (parts.length === 3) {
    const [a, b, c] = parts;
    if (c.length === 4) {
      const year = Number(c);
      const first = Number(a);
      const second = Number(b);

      if (first > 12 && second <= 12) {
        const d = new Date(year, second - 1, first);
        return Number.isNaN(d.getTime()) ? null : d;
      }

      const d = new Date(year, first - 1, second);
      if (!Number.isNaN(d.getTime())) return d;
    }
  }

  const native = new Date(raw);
  if (!Number.isNaN(native.getTime())) {
    return new Date(native.getFullYear(), native.getMonth(), native.getDate());
  }

  return null;
}

function formatDate(date) {
  if (!date || Number.isNaN(date.getTime())) return "";
  const mm = String(date.getMonth() + 1).padStart(2, "0");
  const dd = String(date.getDate()).padStart(2, "0");
  const yyyy = date.getFullYear();
  return `${mm}/${dd}/${yyyy}`;
}

function addMonths(date, months) {
  const d = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  d.setMonth(d.getMonth() + months);
  return d;
}

function parseNumber(value) {
  if (value === null || value === undefined || value === "") return 0;
  if (typeof value === "number") return value;

  const raw = String(value).trim();
  if (!raw) return 0;

  const normalized = raw.replace(/\s/g, "").replace(/[^0-9,.-]/g, "");

  if (!normalized) return 0;

  const hasComma = normalized.includes(",");
  const hasDot = normalized.includes(".");

  if (hasComma && hasDot) {
    const lastComma = normalized.lastIndexOf(",");
    const lastDot = normalized.lastIndexOf(".");
    if (lastComma > lastDot) {
      const value2 = normalized.replace(/\./g, "").replace(",", ".");
      return Number(value2) || 0;
    }
    const value2 = normalized.replace(/,/g, "");
    return Number(value2) || 0;
  }

  if (hasComma && !hasDot) {
    const parts = normalized.split(",");
    if (parts.length > 1 && parts[parts.length - 1].length <= 2) {
      return Number(normalized.replace(",", ".")) || 0;
    }
    return Number(normalized.replace(/,/g, "")) || 0;
  }

  if (!hasComma && hasDot) {
    const parts = normalized.split(".");
    if (parts.length > 1 && parts[parts.length - 1].length <= 2) {
      return Number(normalized) || 0;
    }
    return Number(normalized.replace(/\./g, "")) || 0;
  }

  return Number(normalized) || 0;
}

function money(value) {
  return new Intl.NumberFormat("vi-VN", {
    style: "currency",
    currency: "VND",
    maximumFractionDigits: 0,
  }).format(Number(value || 0));
}

function percent(value) {
  return `${((Number(value || 0)) * 100).toFixed(1)}%`;
}

function downloadWorkbook({ summaryByStore, summaryByProduct, validOrderLines, invalidLines, fileName }) {
  const wb = XLSX.utils.book_new();

  const wsSummaryStore = XLSX.utils.json_to_sheet(summaryByStore);
  const wsSummaryProduct = XLSX.utils.json_to_sheet(summaryByProduct);
  const wsDetails = XLSX.utils.json_to_sheet(validOrderLines);
  const wsInvalid = XLSX.utils.json_to_sheet(invalidLines);

  XLSX.utils.book_append_sheet(wb, wsSummaryStore, "TONG_HOP_NHA_THUOC");
  XLSX.utils.book_append_sheet(wb, wsSummaryProduct, "TONG_HOP_THEO_SP");
  XLSX.utils.book_append_sheet(wb, wsDetails, "CHI_TIET_HOP_LE");
  XLSX.utils.book_append_sheet(wb, wsInvalid, "DON_HANG_LOAI_BO");

  XLSX.writeFile(wb, fileName);
}

function FileCard({ title, description, fileState, onUpload, accent }) {
  return (
    <div className="rounded-3xl border border-slate-200 bg-white p-5 shadow-sm">
      <div className="mb-4 flex items-start justify-between gap-3">
        <div>
          <div className="flex items-center gap-2">
            <div className={`rounded-2xl p-2 ${accent}`}>
              <FileSpreadsheet className="h-5 w-5 text-white" />
            </div>
            <h3 className="text-lg font-semibold text-slate-900">{title}</h3>
          </div>
          <p className="mt-2 text-sm text-slate-500">{description}</p>
        </div>
      </div>

      <label className="flex cursor-pointer items-center justify-center rounded-2xl border border-dashed border-slate-300 bg-slate-50 px-4 py-6 text-center transition hover:border-slate-400 hover:bg-slate-100">
        <input type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={onUpload} />
        <div>
          <Upload className="mx-auto mb-2 h-6 w-6 text-slate-500" />
          <div className="text-sm font-medium text-slate-700">Tải file Excel</div>
          <div className="mt-1 text-xs text-slate-500">Hỗ trợ .xlsx, .xls, .csv</div>
        </div>
      </label>

      {fileState.fileName ? (
        <div className="mt-4 rounded-2xl border border-emerald-200 bg-emerald-50 p-3 text-sm text-emerald-800">
          <div className="font-medium">Đã nạp file: {fileState.fileName}</div>
          <div className="mt-1 text-xs text-emerald-700">
            {fileState.rows.length.toLocaleString("vi-VN")} dòng dữ liệu • {fileState.headers.length} cột
          </div>
        </div>
      ) : null}
    </div>
  );
}

function MappingCard({ title, headers, mapping, requiredFields, onChange }) {
  if (!headers.length) {
    return (
      <div className="rounded-3xl border border-slate-200 bg-white p-5 shadow-sm">
        <h3 className="text-base font-semibold text-slate-900">{title}</h3>
        <p className="mt-2 text-sm text-slate-500">Hãy tải file trước để app tự nhận diện và cho phép map cột.</p>
      </div>
    );
  }

  return (
    <div className="rounded-3xl border border-slate-200 bg-white p-5 shadow-sm">
      <h3 className="text-base font-semibold text-slate-900">{title}</h3>
      <p className="mt-2 text-sm text-slate-500">App đã tự nhận diện cột. Bạn có thể chỉnh lại nếu header thực tế khác tên chuẩn.</p>

      <div className="mt-4 grid gap-4 md:grid-cols-2">
        {requiredFields.map((field) => (
          <div key={field.key}>
            <label className="mb-1 block text-sm font-medium text-slate-700">{field.label}</label>
            <select
              value={mapping[field.key] || ""}
              onChange={(e) => onChange(field.key, e.target.value)}
              className="w-full rounded-2xl border border-slate-300 bg-white px-3 py-2 text-sm outline-none transition focus:border-slate-500"
            >
              <option value="">-- Chọn cột dữ liệu --</option>
              {headers.map((header) => (
                <option key={header} value={header}>
                  {header}
                </option>
              ))}
            </select>
          </div>
        ))}
      </div>
    </div>
  );
}

function PreviewTable({ title, rows }) {
  if (!rows?.length) {
    return (
      <div className="rounded-3xl border border-slate-200 bg-white p-5 shadow-sm">
        <h3 className="text-base font-semibold text-slate-900">{title}</h3>
        <p className="mt-2 text-sm text-slate-500">Chưa có dữ liệu để xem trước.</p>
      </div>
    );
  }

  const columns = Object.keys(rows[0]);

  return (
    <div className="rounded-3xl border border-slate-200 bg-white p-5 shadow-sm">
      <div className="mb-3 flex items-center justify-between gap-3">
        <h3 className="text-base font-semibold text-slate-900">{title}</h3>
        <span className="rounded-full bg-slate-100 px-3 py-1 text-xs text-slate-600">Hiển thị {Math.min(rows.length, SHEET_PREVIEW_LIMIT)} dòng đầu</span>
      </div>

      <div className="overflow-auto rounded-2xl border border-slate-200">
        <table className="min-w-full text-sm">
          <thead className="bg-slate-50 text-slate-700">
            <tr>
              {columns.map((column) => (
                <th key={column} className="whitespace-nowrap border-b border-slate-200 px-3 py-2 text-left font-medium">
                  {column}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.slice(0, SHEET_PREVIEW_LIMIT).map((row, idx) => (
              <tr key={idx} className="odd:bg-white even:bg-slate-50/50">
                {columns.map((column) => (
                  <td key={column} className="whitespace-nowrap border-b border-slate-100 px-3 py-2 text-slate-700">
                    {String(row[column] ?? "")}
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

export default function ContractReconciliationApp() {
  const [contractDurationMonths, setContractDurationMonths] = useState(3);

  const [ordersFile, setOrdersFile] = useState({ fileName: "", rows: [], headers: [], mapping: {} });
  const [productsFile, setProductsFile] = useState({ fileName: "", rows: [], headers: [], mapping: {} });
  const [contractsFile, setContractsFile] = useState({ fileName: "", rows: [], headers: [], mapping: {} });

  const [error, setError] = useState("");
  const [lastRunAt, setLastRunAt] = useState("");

  async function readWorkbook(file, requiredFields) {
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, {
      type: "array",
      cellDates: false,
      raw: true,
    });

    const firstSheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[firstSheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, {
      defval: "",
      raw: true,
    });

    const headers = rows.length ? Object.keys(rows[0]) : [];
    const mapping = autoMapHeaders(headers, requiredFields);

    return {
      fileName: file.name,
      rows,
      headers,
      mapping,
    };
  }

  async function handleUpload(kind, e) {
    const file = e.target.files?.[0];
    if (!file) return;

    setError("");

    try {
      if (kind === "orders") {
        setOrdersFile(await readWorkbook(file, REQUIRED_FIELDS.orders));
      }
      if (kind === "products") {
        setProductsFile(await readWorkbook(file, REQUIRED_FIELDS.products));
      }
      if (kind === "contracts") {
        setContractsFile(await readWorkbook(file, REQUIRED_FIELDS.contracts));
      }
    } catch (err) {
      console.error(err);
      setError(`Không thể đọc file ${file.name}. Hãy kiểm tra định dạng Excel và thử lại.`);
    }
  }

  function updateMapping(kind, fieldKey, headerName) {
    const updater = (prev) => ({
      ...prev,
      mapping: {
        ...prev.mapping,
        [fieldKey]: headerName,
      },
    });

    if (kind === "orders") setOrdersFile(updater);
    if (kind === "products") setProductsFile(updater);
    if (kind === "contracts") setContractsFile(updater);
  }

  const validation = useMemo(() => {
    const messages = [];

    const check = (fileState, requiredFields, fileLabel) => {
      if (!fileState.rows.length) {
        messages.push(`Chưa tải ${fileLabel}.`);
        return;
      }
      requiredFields.forEach((field) => {
        if (!fileState.mapping[field.key]) {
          messages.push(`${fileLabel} chưa map cột: ${field.label}.`);
        }
      });
    };

    check(ordersFile, REQUIRED_FIELDS.orders, "file đơn hàng");
    check(productsFile, REQUIRED_FIELDS.products, "file danh mục sản phẩm");
    check(contractsFile, REQUIRED_FIELDS.contracts, "file hợp đồng");

    return messages;
  }, [ordersFile, productsFile, contractsFile]);

  const processed = useMemo(() => {
    if (validation.length) return null;

    try {
      const productSet = new Set();
      const productNameMap = new Map();

      productsFile.rows.forEach((row) => {
        const code = String(row[productsFile.mapping.productCode] ?? "").trim();
        const name = String(row[productsFile.mapping.productName] ?? "").trim();
        if (!code) return;
        productSet.add(code);
        if (name) productNameMap.set(code, name);
      });

      const contractsByPharmacy = new Map();
      const contractSeedList = [];

      contractsFile.rows.forEach((row, index) => {
        const contractDate = parseFlexibleDate(row[contractsFile.mapping.contractDate]);
        const pharmacyCode = String(row[contractsFile.mapping.pharmacyCode] ?? "").trim();
        const pharmacyName = String(row[contractsFile.mapping.pharmacyName] ?? "").trim();
        const commitment = parseNumber(row[contractsFile.mapping.commitment]);

        if (!contractDate || !pharmacyCode) return;

        const endExclusive = contractDurationMonths > 0 ? addMonths(contractDate, contractDurationMonths) : null;
        const contractId = `${pharmacyCode}__${formatDate(contractDate)}__${index}`;

        const contractRecord = {
          contractId,
          contractDate,
          endExclusive,
          pharmacyCode,
          pharmacyName,
          commitment,
        };

        if (!contractsByPharmacy.has(pharmacyCode)) contractsByPharmacy.set(pharmacyCode, []);
        contractsByPharmacy.get(pharmacyCode).push(contractRecord);
        contractSeedList.push(contractRecord);
      });

      contractsByPharmacy.forEach((list) => {
        list.sort((a, b) => a.contractDate - b.contractDate);
      });

      const summaryByStoreMap = new Map();
      const summaryByProductMap = new Map();
      const validOrderLines = [];
      const invalidLines = [];

      contractSeedList.forEach((contract) => {
        summaryByStoreMap.set(contract.contractId, {
          "Ngày ký hợp đồng": formatDate(contract.contractDate),
          "Mã nhà thuốc": contract.pharmacyCode,
          "Tên nhà thuốc": contract.pharmacyName,
          "Doanh số": 0,
          "Mức cam kết hợp đồng": contract.commitment,
          "Tỉ lệ đạt": 0,
          "Trạng thái": "Chưa phát sinh",
        });
      });

      const counters = {
        totalLines: ordersFile.rows.length,
        validLines: 0,
        excludedProduct: 0,
        excludedNoContract: 0,
        excludedBeforeOrOutsideWindow: 0,
        invalidData: 0,
      };

      ordersFile.rows.forEach((row, idx) => {
        const orderDate = parseFlexibleDate(row[ordersFile.mapping.orderDate]);
        const pharmacyCode = String(row[ordersFile.mapping.pharmacyCode] ?? "").trim();
        const pharmacyName = String(row[ordersFile.mapping.pharmacyName] ?? "").trim();
        const productCode = String(row[ordersFile.mapping.productCode] ?? "").trim();
        const productName = String(row[ordersFile.mapping.productName] ?? "").trim() || productNameMap.get(productCode) || "";
        const sales = parseNumber(row[ordersFile.mapping.sales]);

        if (!orderDate || !pharmacyCode || !productCode) {
          counters.invalidData += 1;
          invalidLines.push({
            "Dòng nguồn": idx + 2,
            "Lý do loại": "Thiếu ngày đơn hàng / mã nhà thuốc / mã sản phẩm",
            "Ngày đơn hàng": row[ordersFile.mapping.orderDate] ?? "",
            "Mã nhà thuốc": pharmacyCode,
            "Tên nhà thuốc": pharmacyName,
            "Mã sản phẩm": productCode,
            "Tên sản phẩm": productName,
            "Doanh số": sales,
          });
          return;
        }

        if (!productSet.has(productCode)) {
          counters.excludedProduct += 1;
          invalidLines.push({
            "Dòng nguồn": idx + 2,
            "Lý do loại": "Mã sản phẩm không thuộc danh mục sản phẩm tham gia hợp đồng",
            "Ngày đơn hàng": formatDate(orderDate),
            "Mã nhà thuốc": pharmacyCode,
            "Tên nhà thuốc": pharmacyName,
            "Mã sản phẩm": productCode,
            "Tên sản phẩm": productName,
            "Doanh số": sales,
          });
          return;
        }

        const contractList = contractsByPharmacy.get(pharmacyCode);
        if (!contractList?.length) {
          counters.excludedNoContract += 1;
          invalidLines.push({
            "Dòng nguồn": idx + 2,
            "Lý do loại": "Nhà thuốc không có trong file hợp đồng",
            "Ngày đơn hàng": formatDate(orderDate),
            "Mã nhà thuốc": pharmacyCode,
            "Tên nhà thuốc": pharmacyName,
            "Mã sản phẩm": productCode,
            "Tên sản phẩm": productName,
            "Doanh số": sales,
          });
          return;
        }

        const eligibleContracts = contractList.filter((contract) => {
          const meetsStart = orderDate >= contract.contractDate;
          const meetsEnd = !contract.endExclusive || orderDate < contract.endExclusive;
          return meetsStart && meetsEnd;
        });

        if (!eligibleContracts.length) {
          counters.excludedBeforeOrOutsideWindow += 1;
          invalidLines.push({
            "Dòng nguồn": idx + 2,
            "Lý do loại": contractDurationMonths > 0
              ? `Đơn hàng nằm ngoài thời gian hiệu lực hợp đồng ${contractDurationMonths} tháng`
              : "Đơn hàng phát sinh trước ngày ký hợp đồng",
            "Ngày đơn hàng": formatDate(orderDate),
            "Mã nhà thuốc": pharmacyCode,
            "Tên nhà thuốc": pharmacyName,
            "Mã sản phẩm": productCode,
            "Tên sản phẩm": productName,
            "Doanh số": sales,
          });
          return;
        }

        const chosenContract = eligibleContracts.sort((a, b) => b.contractDate - a.contractDate)[0];

        counters.validLines += 1;

        const storeSummary = summaryByStoreMap.get(chosenContract.contractId);
        storeSummary["Doanh số"] += sales;
        storeSummary["Tỉ lệ đạt"] = chosenContract.commitment > 0 ? storeSummary["Doanh số"] / chosenContract.commitment : 0;
        storeSummary["Trạng thái"] = storeSummary["Tỉ lệ đạt"] >= 1 ? "Đạt cam kết" : "Chưa đạt";

        const productSummaryKey = `${chosenContract.contractId}__${productCode}`;
        if (!summaryByProductMap.has(productSummaryKey)) {
          summaryByProductMap.set(productSummaryKey, {
            "Ngày ký hợp đồng": formatDate(chosenContract.contractDate),
            "Mã nhà thuốc": chosenContract.pharmacyCode,
            "Tên nhà thuốc": chosenContract.pharmacyName,
            "Mã sản phẩm": productCode,
            "Tên sản phẩm": productName,
            "Doanh số": 0,
            "Mức cam kết hợp đồng": chosenContract.commitment,
            "Tỉ lệ đạt": 0,
          });
        }

        const productSummary = summaryByProductMap.get(productSummaryKey);
        productSummary["Doanh số"] += sales;
        productSummary["Tỉ lệ đạt"] = chosenContract.commitment > 0 ? productSummary["Doanh số"] / chosenContract.commitment : 0;

        validOrderLines.push({
          "Ngày ký hợp đồng": formatDate(chosenContract.contractDate),
          "Ngày đơn hàng": formatDate(orderDate),
          "Mã nhà thuốc": chosenContract.pharmacyCode,
          "Tên nhà thuốc": chosenContract.pharmacyName || pharmacyName,
          "Mã sản phẩm": productCode,
          "Tên sản phẩm": productName,
          "Doanh số": sales,
          "Mức cam kết hợp đồng": chosenContract.commitment,
          "Tỉ lệ đạt": chosenContract.commitment > 0 ? sales / chosenContract.commitment : 0,
        });
      });

      const summaryByStore = Array.from(summaryByStoreMap.values())
        .map((row) => ({
          ...row,
          "Tỉ lệ đạt": Number((row["Tỉ lệ đạt"] || 0).toFixed(4)),
        }))
        .sort((a, b) => a["Mã nhà thuốc"].localeCompare(b["Mã nhà thuốc"]));

      const summaryByProduct = Array.from(summaryByProductMap.values())
        .map((row) => ({
          ...row,
          "Tỉ lệ đạt": Number((row["Tỉ lệ đạt"] || 0).toFixed(4)),
        }))
        .sort((a, b) => {
          const c1 = a["Mã nhà thuốc"].localeCompare(b["Mã nhà thuốc"]);
          if (c1 !== 0) return c1;
          return a["Mã sản phẩm"].localeCompare(b["Mã sản phẩm"]);
        });

      const validOrderLinesFormatted = validOrderLines.map((row) => ({
        ...row,
        "Tỉ lệ đạt": Number((row["Tỉ lệ đạt"] || 0).toFixed(4)),
      }));

      return {
        counters,
        summaryByStore,
        summaryByProduct,
        validOrderLines: validOrderLinesFormatted,
        invalidLines,
      };
    } catch (err) {
      console.error(err);
      return {
        error: "Có lỗi khi xử lý dữ liệu. Hãy kiểm tra lại mapping cột, kiểu dữ liệu ngày và số tiền trong file Excel.",
      };
    }
  }, [ordersFile, productsFile, contractsFile, validation, contractDurationMonths]);

  const topSummaryPreview = useMemo(() => {
    if (!processed?.summaryByStore?.length) return [];
    return processed.summaryByStore.slice(0, SHEET_PREVIEW_LIMIT).map((row) => ({
      ...row,
      "Doanh số": money(row["Doanh số"]),
      "Mức cam kết hợp đồng": money(row["Mức cam kết hợp đồng"]),
      "Tỉ lệ đạt": percent(row["Tỉ lệ đạt"]),
    }));
  }, [processed]);

  const topProductPreview = useMemo(() => {
    if (!processed?.summaryByProduct?.length) return [];
    return processed.summaryByProduct.slice(0, SHEET_PREVIEW_LIMIT).map((row) => ({
      ...row,
      "Doanh số": money(row["Doanh số"]),
      "Mức cam kết hợp đồng": money(row["Mức cam kết hợp đồng"]),
      "Tỉ lệ đạt": percent(row["Tỉ lệ đạt"]),
    }));
  }, [processed]);

  const topInvalidPreview = useMemo(() => {
    if (!processed?.invalidLines?.length) return [];
    return processed.invalidLines.slice(0, SHEET_PREVIEW_LIMIT).map((row) => ({
      ...row,
      "Doanh số": money(row["Doanh số"]),
    }));
  }, [processed]);

  function handleExport() {
    if (!processed || processed.error) return;

    downloadWorkbook({
      summaryByStore: processed.summaryByStore.map((row) => ({
        ...row,
        "Doanh số": Number(row["Doanh số"].toFixed(0)),
        "Mức cam kết hợp đồng": Number(row["Mức cam kết hợp đồng"].toFixed(0)),
      })),
      summaryByProduct: processed.summaryByProduct.map((row) => ({
        ...row,
        "Doanh số": Number(row["Doanh số"].toFixed(0)),
        "Mức cam kết hợp đồng": Number(row["Mức cam kết hợp đồng"].toFixed(0)),
      })),
      validOrderLines: processed.validOrderLines.map((row) => ({
        ...row,
        "Doanh số": Number(row["Doanh số"].toFixed(0)),
        "Mức cam kết hợp đồng": Number(row["Mức cam kết hợp đồng"].toFixed(0)),
      })),
      invalidLines: processed.invalidLines.map((row) => ({
        ...row,
        "Doanh số": Number(parseNumber(row["Doanh số"]).toFixed(0)),
      })),
      fileName: `doi-soat-hop-dong-tich-luy-3-thang_${new Date().toISOString().slice(0, 10)}.xlsx`,
    });

    setLastRunAt(new Date().toLocaleString("vi-VN"));
  }

  function handleReset() {
    setOrdersFile({ fileName: "", rows: [], headers: [], mapping: {} });
    setProductsFile({ fileName: "", rows: [], headers: [], mapping: {} });
    setContractsFile({ fileName: "", rows: [], headers: [], mapping: {} });
    setError("");
    setLastRunAt("");
  }

  return (
    <div className="min-h-screen bg-slate-100 text-slate-900">
      <div className="mx-auto max-w-7xl px-4 py-8 md:px-6 lg:px-8">
        <div className="mb-8 rounded-[28px] bg-gradient-to-r from-slate-900 via-slate-800 to-slate-700 p-6 text-white shadow-xl">
          <div className="grid gap-6 lg:grid-cols-[1.5fr_0.8fr]">
            <div>
              <div className="mb-3 inline-flex items-center gap-2 rounded-full border border-white/20 bg-white/10 px-3 py-1 text-xs font-medium uppercase tracking-wide text-slate-100">
                <ShieldCheck className="h-4 w-4" />
                Web app đối soát nội bộ
              </div>
              <h1 className="text-2xl font-bold tracking-tight md:text-4xl">ĐỐI SOÁT HỢP ĐỒNG TÍCH LŨY 3 THÁNG</h1>
              <p className="mt-3 max-w-3xl text-sm leading-6 text-slate-200 md:text-base">
                Ứng dụng giúp đối chiếu dữ liệu đơn hàng, danh mục sản phẩm và hợp đồng nhà thuốc để xác định chính xác doanh số tích lũy hợp lệ theo từng thời điểm ký hợp đồng trong năm 2026.
              </p>
            </div>

            <div className="rounded-3xl border border-white/10 bg-white/10 p-5 backdrop-blur">
              <div className="text-sm font-semibold text-white">Thiết lập nhanh</div>
              <div className="mt-3">
                <label className="mb-1 block text-sm text-slate-200">Thời gian hiệu lực hợp đồng</label>
                <div className="flex items-center gap-3">
                  <input
                    type="number"
                    min={0}
                    value={contractDurationMonths}
                    onChange={(e) => setContractDurationMonths(Math.max(0, Number(e.target.value) || 0))}
                    className="w-28 rounded-2xl border border-white/20 bg-white/10 px-3 py-2 text-white outline-none placeholder:text-slate-300 focus:border-white/40"
                  />
                  <span className="text-sm text-slate-200">tháng (mặc định = 3)</span>
                </div>
                <div className="mt-2 text-xs text-slate-300">
                  Đặt = 0 nếu bạn muốn tính từ ngày ký hợp đồng trở đi mà không giới hạn số tháng.
                </div>
              </div>
            </div>
          </div>
        </div>

        <div className="mb-8 grid gap-4 lg:grid-cols-3">
          <FileCard
            title="1. File dữ liệu đơn hàng"
            description="Bao gồm ngày đơn hàng, mã nhà thuốc, tên nhà thuốc, mã sản phẩm, tên sản phẩm, doanh số."
            fileState={ordersFile}
            onUpload={(e) => handleUpload("orders", e)}
            accent="bg-blue-600"
          />
          <FileCard
            title="2. File danh mục sản phẩm"
            description="Danh mục chuẩn các mã sản phẩm được phép tham gia tích lũy hợp đồng."
            fileState={productsFile}
            onUpload={(e) => handleUpload("products", e)}
            accent="bg-emerald-600"
          />
          <FileCard
            title="3. File hợp đồng nhà thuốc"
            description="Bao gồm ngày đăng ký hợp đồng, mã nhà thuốc, tên nhà thuốc và mức doanh số cam kết."
            fileState={contractsFile}
            onUpload={(e) => handleUpload("contracts", e)}
            accent="bg-violet-600"
          />
        </div>

        <div className="mb-8 grid gap-4 lg:grid-cols-3">
          <MappingCard
            title="Map cột file đơn hàng"
            headers={ordersFile.headers}
            mapping={ordersFile.mapping}
            requiredFields={REQUIRED_FIELDS.orders}
            onChange={(fieldKey, value) => updateMapping("orders", fieldKey, value)}
          />
          <MappingCard
            title="Map cột file danh mục sản phẩm"
            headers={productsFile.headers}
            mapping={productsFile.mapping}
            requiredFields={REQUIRED_FIELDS.products}
            onChange={(fieldKey, value) => updateMapping("products", fieldKey, value)}
          />
          <MappingCard
            title="Map cột file hợp đồng"
            headers={contractsFile.headers}
            mapping={contractsFile.mapping}
            requiredFields={REQUIRED_FIELDS.contracts}
            onChange={(fieldKey, value) => updateMapping("contracts", fieldKey, value)}
          />
        </div>

        {error ? (
          <div className="mb-6 rounded-3xl border border-rose-200 bg-rose-50 p-4 text-sm text-rose-700">
            <div className="flex items-center gap-2 font-semibold">
              <AlertCircle className="h-4 w-4" />
              Lỗi đọc file
            </div>
            <div className="mt-1">{error}</div>
          </div>
        ) : null}

        {validation.length ? (
          <div className="mb-6 rounded-3xl border border-amber-200 bg-amber-50 p-4 text-sm text-amber-800">
            <div className="flex items-center gap-2 font-semibold">
              <AlertCircle className="h-4 w-4" />
              Cần hoàn tất các bước trước khi đối soát
            </div>
            <ul className="mt-2 space-y-1 pl-5 text-sm list-disc">
              {validation.map((message, index) => (
                <li key={index}>{message}</li>
              ))}
            </ul>
          </div>
        ) : null}

        {processed?.error ? (
          <div className="mb-6 rounded-3xl border border-rose-200 bg-rose-50 p-4 text-sm text-rose-700">
            <div className="flex items-center gap-2 font-semibold">
              <AlertCircle className="h-4 w-4" />
              Không thể xử lý dữ liệu
            </div>
            <div className="mt-1">{processed.error}</div>
          </div>
        ) : null}

        <div className="mb-8 flex flex-wrap items-center gap-3">
          <button
            onClick={handleExport}
            disabled={!processed || !!processed?.error || !!validation.length}
            className="inline-flex items-center gap-2 rounded-2xl bg-slate-900 px-4 py-3 text-sm font-medium text-white transition hover:bg-slate-800 disabled:cursor-not-allowed disabled:bg-slate-400"
          >
            <Download className="h-4 w-4" />
            Xuất file Excel đối soát
          </button>
          <button
            onClick={handleReset}
            className="inline-flex items-center gap-2 rounded-2xl border border-slate-300 bg-white px-4 py-3 text-sm font-medium text-slate-700 transition hover:bg-slate-50"
          >
            <RefreshCw className="h-4 w-4" />
            Làm mới dữ liệu
          </button>
          {lastRunAt ? <span className="text-sm text-slate-500">Lần xuất gần nhất: {lastRunAt}</span> : null}
        </div>

        {processed ? (
          <>
            <div className="mb-8 grid gap-4 md:grid-cols-2 xl:grid-cols-5">
              <div className="rounded-3xl border border-slate-200 bg-white p-5 shadow-sm">
                <div className="text-sm text-slate-500">Tổng dòng đơn hàng</div>
                <div className="mt-2 text-3xl font-bold text-slate-900">{processed.counters.totalLines.toLocaleString("vi-VN")}</div>
              </div>
              <div className="rounded-3xl border border-emerald-200 bg-emerald-50 p-5 shadow-sm">
                <div className="text-sm text-emerald-700">Dòng hợp lệ</div>
                <div className="mt-2 text-3xl font-bold text-emerald-800">{processed.counters.validLines.toLocaleString("vi-VN")}</div>
              </div>
              <div className="rounded-3xl border border-amber-200 bg-amber-50 p-5 shadow-sm">
                <div className="text-sm text-amber-700">Loại do sai sản phẩm</div>
                <div className="mt-2 text-3xl font-bold text-amber-800">{processed.counters.excludedProduct.toLocaleString("vi-VN")}</div>
              </div>
              <div className="rounded-3xl border border-rose-200 bg-rose-50 p-5 shadow-sm">
                <div className="text-sm text-rose-700">Loại do không có hợp đồng</div>
                <div className="mt-2 text-3xl font-bold text-rose-800">{processed.counters.excludedNoContract.toLocaleString("vi-VN")}</div>
              </div>
              <div className="rounded-3xl border border-violet-200 bg-violet-50 p-5 shadow-sm">
                <div className="text-sm text-violet-700">Loại do ngày không hợp lệ</div>
                <div className="mt-2 text-3xl font-bold text-violet-800">{(processed.counters.excludedBeforeOrOutsideWindow + processed.counters.invalidData).toLocaleString("vi-VN")}</div>
              </div>
            </div>

            <div className="mb-8 grid gap-4 lg:grid-cols-2">
              <div className="rounded-3xl border border-slate-200 bg-white p-5 shadow-sm">
                <div className="mb-3 flex items-center gap-2 text-slate-900">
                  <Database className="h-5 w-5" />
                  <h3 className="text-base font-semibold">Logic đối soát đang áp dụng</h3>
                </div>
                <ul className="space-y-2 text-sm leading-6 text-slate-600 list-disc pl-5">
                  <li>Chỉ tính đơn hàng có mã sản phẩm nằm trong file danh mục sản phẩm tham gia hợp đồng.</li>
                  <li>Chỉ tính đơn hàng của nhà thuốc xuất hiện trong file hợp đồng.</li>
                  <li>Chỉ tính đơn hàng phát sinh từ ngày ký hợp đồng trở đi.</li>
                  <li>
                    {contractDurationMonths > 0
                      ? `Mỗi hợp đồng được tính trong cửa sổ hiệu lực ${contractDurationMonths} tháng kể từ ngày ký.`
                      : "Hợp đồng được tính từ ngày ký trở đi và không giới hạn số tháng."}
                  </li>
                  <li>Nếu cùng một nhà thuốc có nhiều hợp đồng, app sẽ gán đơn hàng vào hợp đồng hợp lệ gần nhất theo ngày ký.</li>
                </ul>
              </div>

              <div className="rounded-3xl border border-slate-200 bg-white p-5 shadow-sm">
                <div className="mb-3 flex items-center gap-2 text-slate-900">
                  <CheckCircle2 className="h-5 w-5" />
                  <h3 className="text-base font-semibold">File Excel xuất ra gồm 4 sheet</h3>
                </div>
                <ul className="space-y-2 text-sm leading-6 text-slate-600 list-disc pl-5">
                  <li><strong>TONG_HOP_NHA_THUOC</strong>: Tổng doanh số hợp lệ theo từng nhà thuốc / từng hợp đồng.</li>
                  <li><strong>TONG_HOP_THEO_SP</strong>: Tổng hợp theo nhà thuốc + mã sản phẩm, đúng cấu trúc bạn cần để đối chiếu.</li>
                  <li><strong>CHI_TIET_HOP_LE</strong>: Danh sách từng dòng đơn hàng được chấp nhận.</li>
                  <li><strong>DON_HANG_LOAI_BO</strong>: Danh sách các dòng bị loại và lý do loại để kiểm tra lại dữ liệu.</li>
                </ul>
              </div>
            </div>

            <div className="mb-8 grid gap-4 lg:grid-cols-2">
              <PreviewTable title="Xem trước tổng hợp nhà thuốc" rows={topSummaryPreview} />
              <PreviewTable title="Xem trước tổng hợp theo sản phẩm" rows={topProductPreview} />
            </div>

            <PreviewTable title="Xem trước các dòng bị loại" rows={topInvalidPreview} />
          </>
        ) : null}

        <div className="mt-8 rounded-3xl border border-slate-200 bg-white p-5 shadow-sm">
          <h3 className="text-base font-semibold text-slate-900">Lưu ý dữ liệu đầu vào</h3>
          <div className="mt-3 grid gap-4 md:grid-cols-3">
            <div className="rounded-2xl bg-slate-50 p-4 text-sm text-slate-600">
              <div className="font-semibold text-slate-800">Ngày tháng</div>
              <div className="mt-1">Ưu tiên định dạng MM/DD/YYYY. App vẫn cố gắng đọc các định dạng ngày phổ biến và Excel date serial.</div>
            </div>
            <div className="rounded-2xl bg-slate-50 p-4 text-sm text-slate-600">
              <div className="font-semibold text-slate-800">Doanh số / cam kết</div>
              <div className="mt-1">Có thể là số thường hoặc chuỗi có dấu phân cách hàng nghìn. App sẽ tự chuẩn hóa khi tính toán.</div>
            </div>
            <div className="rounded-2xl bg-slate-50 p-4 text-sm text-slate-600">
              <div className="font-semibold text-slate-800">Tên cột không chuẩn</div>
              <div className="mt-1">Nếu file thực tế khác tên chuẩn, bạn chỉ cần chỉnh lại phần map cột phía trên rồi xuất kết quả.</div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}
