// ==========================================
// 1. ربط عناصر واجهة المستخدم (DOM Elements)
// ==========================================
const fileInput = document.getElementById("fileInput"); // حقل اختيار الملف
const convertBtn = document.getElementById("convertBtn"); // زر بدء التحويل
const downloadBtn = document.getElementById("downloadBtn"); // زر التنزيل (معطل افتراضياً)
const tableBody = document.querySelector("#previewTable tbody"); // جسم الجدول للعرض
const stats = document.getElementById("stats"); // لعرض عدد الأسطر الناتجة
let barcode = document.getElementById("startBarcode").value || "TBJ123";

// متغير عام لتخزين البيانات المعالجة لتكون جاهزة للتنزيل لاحقاً
let PhenixData = [];

// ==========================================
// 2. حدث النقر على زر التحويل
// ==========================================
convertBtn.onclick = async () => {
   const file = fileInput.files[0];
   if (!file) {
      alert("الرجاء اختيار ملف Excel أولاً!");
      return;
   }

   const rawData = await readExcel(file);
   console.log("✅ تم قراءة البيانات الخام:", rawData);

   await loadColorIds();
   await loadClassItems();

   // ✅ اقرأ القيم هنا (بعد اختيار المستخدم)
   const transformType = document.getElementById("transformType").value;
   barcode = document.getElementById("startBarcode").value || "TBJ123";

   console.log("🔧 Transform Type:", transformType);
   console.log("🏷️ Start Barcode:", barcode);

   if (transformType === "bags") {
      PhenixData = transformBags(rawData);
   } else {
      PhenixData = transform(rawData);
   }

   console.log("🚀 البيانات بعد المعالجة (PhenixData):", PhenixData);

   renderTable(PhenixData);

   stats.innerText = `عدد الأسطر الناتجة: ${PhenixData.length}`;
   downloadBtn.disabled = false;
};

// ==========================================
// 3. حدث النقر على زر التنزيل
// ==========================================
downloadBtn.onclick = () => {
   // استدعاء دالة التصدير الموجودة في excel.js
   downloadExcel(PhenixData);
};

// ==========================================
// 4. دالة التحويل الأساسية (Logic Core)
// هذه الدالة تحول الجدول المحوري (Pivot-like) إلى جدول بيانات عادي
// ==========================================
function transform(data) {
   console.log("🔁 transform started (FINAL OUTPUT)");

   const result = [];

   let currentItemCode = null;
   let currentClassCode = null;
   let lastOutputItemCode = null;
   let currentCTNS = 0;
   let currentCTNSQty = 0;
   let currentTTL = 0;
   let currentPrice = 0;
   let currentAmount = 0;
   // let barcode = "TBJ123";

   // ------------------------------------------------
   // 1) اكتشاف المقاسات من صف QTY
   // ------------------------------------------------
   const sizeMap = {};
   let headerRowFound = false;

   for (const row of data) {
      if (row.__EMPTY_3 === "QTY") {
         Object.keys(row).forEach((key) => {
            const val = row[key];
            if (typeof val === "number" && val > 0) {
               sizeMap[key] = val;
            }
         });
         headerRowFound = true;
         break;
      }
   }

   if (!headerRowFound) {
      console.error("❌ QTY row not found");
      return [];
   }

   // ------------------------------------------------
   // 2) المرور على البيانات
   // ------------------------------------------------
   data.forEach((row) => {
      const itemCell = row.__EMPTY; // ITEM NO
      const colorCell = row.__EMPTY_1; // COLOUR

      // صف صنف جديد
      if (typeof itemCell === "string" && itemCell.trim() !== "") {
         currentItemCode = itemCell.trim().replaceAll(/\s/g, "");
         currentClassCode = extractClassCode(currentItemCode);
         barcode = nextCode(barcode);
         currentCTNS = Number(row.__EMPTY_2) || 0;
         currentCTNSQty = Number(row.__EMPTY_3) || 0;
         currentTTL = Number(row.__EMPTY_4) || 0;
         currentPrice = Number(row.__EMPTY_5) || 0;
         currentAmount = Number(row.__EMPTY_6) || 0;
      }

      if (!currentItemCode || !colorCell || typeof colorCell !== "string")
         return;

      const colorName = colorCell.trim();
      const colorId = getColorId(colorName);

      // ------------------------------------------------
      // 3) تفكيك المقاسات + الحسابات
      // ------------------------------------------------
      Object.entries(sizeMap).forEach(([colKey, size]) => {
         const qty = Number(row[colKey]) || 0;
         if (qty <= 0) return;

         const qtyCTNS = qty * currentCTNS;
         const qtyCTNSPrice = qtyCTNS * currentPrice;

         const isFirstRowOfItem = currentItemCode !== lastOutputItemCode;

         result.push({
            PICTURE: "",

            "ITEM NO": currentItemCode,
            ClassCode: currentClassCode,
            Barcode: barcode,
            color: colorName,
            "Id Color": colorId,

            // 👇 هذه الأعمدة فقط في أول سطر
            CTNS: isFirstRowOfItem ? currentCTNS : "",
            "CTNS / QTY": isFirstRowOfItem ? currentCTNSQty : "",
            TTL: currentTTL,
            PRICE: currentPrice,
            AMOUNT: isFirstRowOfItem ? currentAmount : "",

            // 👇 هذه في كل سطر
            size: size,
            quantity: qty,
            "quantity * CTNS": qtyCTNS,
            "quantity * CTNS * PRICE": qtyCTNSPrice,
         });

         // تحديث آخر مادة تمت كتابتها
         lastOutputItemCode = currentItemCode;
      });
   });

   console.log("✅ transform finished");
   console.log("📦 rows:", result.length);

   return result;
}

// ==========================================

// ==========================================
function normalizeColorQuantities(colors, targetTotal) {
   const originalTotal = colors.reduce((s, c) => s + c.qty, 0);
   if (originalTotal === 0) return colors;

   // 1) توزيع نسبي
   let normalized = colors.map((c) => ({
      color: c.color,
      qty: Math.floor((c.qty / originalTotal) * targetTotal),
   }));

   // 2) إصلاح الفرق
   let currentTotal = normalized.reduce((s, c) => s + c.qty, 0);
   let diff = targetTotal - currentTotal;

   let i = 0;
   while (diff !== 0) {
      normalized[i % normalized.length].qty += diff > 0 ? 1 : -1;
      diff += diff > 0 ? -1 : 1;
      i++;
   }

   return normalized;
}
function transformBags(data) {
   console.log("👜 transformBags started");

   const result = [];
   //   let barcode = "TBJ123";

   data.forEach((row, index) => {
      const itemCode = row.__EMPTY; // ITEM NO
      const colorsCell = row.__EMPTY_1; // colors string
      const totalQty = Number(row.__EMPTY_4) || 0; // TOTAL / PCS
      const price = Number(row.__EMPTY_5) || 0; // PRICE

      // Debug دقيق (شغّله لو لزم)
      // console.log(index, itemCode, colorsCell, totalQty, price);

      if (
         typeof itemCode !== "string" ||
         typeof colorsCell !== "string" ||
         totalQty <= 0 ||
         price <= 0
      ) {
         return;
      }
      barcode = nextCode(barcode);
      // ----------------------------------
      // 1) تحليل خلية الألوان
      // ----------------------------------
      const colors = [];
      const regex = /([a-zA-Z\s\-]+)\s*(\d+)/g;
      let match;

      while ((match = regex.exec(colorsCell)) !== null) {
         colors.push({
            color: match[1].trim().toUpperCase(),
            qty: Number(match[2]),
         });
      }

      if (colors.length === 0) return;

      // ----------------------------------
      // 2) توحيد المجموع مع TOTAL / PCS
      // ----------------------------------
      const normalizedColors = normalizeColorQuantities(colors, totalQty);

      // ----------------------------------
      // 3) إخراج الصفوف
      // ----------------------------------
      normalizedColors.forEach((c) => {
         if (c.qty <= 0) return;
         result.push({
            PICTURE: "لا يوجد",
            "ITEM NO": itemCode.trim(),
            ClassCode: "لا يوجد",
            color: c.color,
            "Id Color": getColorId(c.color),
            Barcode: barcode,
            quantity: c.qty,
            PRICE: price,
            AMOUNT: c.qty * price,
         });
      });
   });

   console.log("👜 transformBags finished");
   console.log("📦 rows:", result.length);

   return result;
}

// ==========================================
// 5. دالة رسم الجدول (UI Helper)
// ==========================================
const OUTPUT_COLUMNS = [
   { key: "PICTURE", label: "PICTURE" },
   { key: "ITEM NO", label: "ITEM NO" },
   { key: "ClassCode", label: "ClassCode" },
   { key: "Barcode", label: "Barcode" },
   { key: "color", label: "Color" },
   { key: "Id Color", label: "Color ID" },
   { key: "CTNS", label: "CTNS" },
   { key: "CTNS / QTY", label: "CTNS / QTY" },
   { key: "TTL", label: "TTL" },
   { key: "PRICE", label: "PRICE" },
   { key: "AMOUNT", label: "AMOUNT" },
   { key: "size", label: "Size" },
   { key: "quantity", label: "Qty" },
   { key: "quantity * CTNS", label: "Qty × CTNS" },
   { key: "quantity * CTNS * PRICE", label: "Qty × CTNS × PRICE" },
];

function renderTable(rows) {
   const table = document.getElementById("previewTable");

   // مسح أي محتوى قديم
   table.innerHTML = "";

   if (!rows || rows.length === 0) {
      table.innerHTML = "<tr><td>لا توجد بيانات للعرض</td></tr>";
      return;
   }

   // -----------------------------
   // 1) إنشاء الرأس (thead)
   // -----------------------------
   const thead = document.createElement("thead");
   const headRow = document.createElement("tr");

   OUTPUT_COLUMNS.forEach((col) => {
      const th = document.createElement("th");
      th.textContent = col.label;
      headRow.appendChild(th);
   });

   thead.appendChild(headRow);
   table.appendChild(thead);

   // -----------------------------
   // 2) إنشاء الجسم (tbody)
   // -----------------------------
   const tbody = document.createElement("tbody");

   rows.forEach((row) => {
      const tr = document.createElement("tr");

      OUTPUT_COLUMNS.forEach((col) => {
         const td = document.createElement("td");
         const value = row[col.key];

         td.textContent =
            value === undefined || value === null || value === "" ? "" : value;

         tr.appendChild(td);
      });

      tbody.appendChild(tr);
   });

   table.appendChild(tbody);
}

// ==========================================
// 6. دالة مساعدة لاستخراج اسم الصنف (Utility)
// الهدف: إزالة الأرقام والحرف الأول للحصول على "تصنيف"
// مثال: "ZX3020" -> تصبح "ZX" (حسب المنطق المكتوب) أو حسب الحروف غير الرقمية
// ==========================================
function extractClassCode(itemCode) {
   let classCode = "";
   // ملاحظة: التعامل مع string كـ array
   for (const i in itemCode) {
      // تجاوز الحرف الأول (index 0)
      if (i == 0) {
         continue;
      }
      // إذا كان الحرف ليس رقماً، أضفه للاسم
      if (isNaN(Number(itemCode[i]))) {
         classCode += itemCode[i];
      }
   }
   return classCode;
}

let classItemMap = null;

// تحميل ملف التصنيفات مرة واحدة
async function loadClassItems() {
   if (classItemMap) return classItemMap;

   const response = await fetch("classItemsNames.json");
   const data = await response.json();

   classItemMap = {};

   data.forEach((item) => {
      classItemMap[item.ClassItemCode.toUpperCase()] = item.ClassName;
   });

   console.log("📦 Class Items loaded:", classItemMap);

   return classItemMap;
}

function getItemClass(classCode) {
   if (!classItemMap) {
      console.warn("⚠️ ClassItem map not loaded yet");
      return "";
   }

   const key = classCode.trim().toUpperCase();
   return classItemMap[key] || "";
}

// ==========================================

// ==========================================
let colorIdMap = null;

// تحميل ملف الألوان مرة واحدة
async function loadColorIds() {
   if (colorIdMap) return colorIdMap;

   const response = await fetch("colorsIds.json");
   const data = await response.json();

   colorIdMap = {};

   data.forEach((item) => {
      colorIdMap[item.ColorName.toUpperCase()] = item.ColorId;
   });

   console.log("🎨 Color IDs loaded:", colorIdMap);

   return colorIdMap;
}

function getColorId(colorName) {
   if (!colorIdMap) {
      console.warn("⚠️ ColorId map not loaded yet");
      return "00";
   }

   const key = colorName.trim().toUpperCase();
   return colorIdMap[key] || "";
}

// ==========================================

// ==========================================
function nextCode(code) {
   // فصل الجزء الحرفي عن الرقمي
   let letters = code.slice(0, 3);
   let number = parseInt(code.slice(3), 10);

   number++;

   // إذا تجاوزنا 999
   if (number > 999) {
      number = 0;
      letters = incrementLetters(letters);
   }

   return letters + number.toString().padStart(3, "0");
}

function incrementLetters(str) {
   let chars = str.split("");

   for (let i = chars.length - 1; i >= 0; i--) {
      if (chars[i] !== "Z") {
         chars[i] = String.fromCharCode(chars[i].charCodeAt(0) + 1);
         break;
      } else {
         chars[i] = "A";
      }
   }

   return chars.join("");
}
