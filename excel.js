// ==========================================
// دالة قراءة ملف Excel
// تُرجع Promise لأن قراءة الملفات عملية غير متزامنة (Async)
// ==========================================
function readExcel(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    // 1. عند اكتمال تحميل الملف إلى الذاكرة
    reader.onload = (e) => {
      console.log("📥 تم تحميل الملف إلى المتصفح");

      // تحويل البيانات إلى مصفوفة بايتات (Uint8Array) لتفهمها المكتبة
      const data = new Uint8Array(e.target.result);
      console.log("📦 حجم البيانات (بايت):", data.length);

      // 2. قراءة البيانات بواسطة مكتبة XLSX
      const workbook = XLSX.read(data, { type: "array" });

      // طباعة أسماء الصفحات (Sheets) للتأكد
      console.log("📄 الصفحات الموجودة:", workbook.SheetNames);

      // اختيار الصفحة الأولى دائماً
      const sheet = workbook.Sheets[workbook.SheetNames[0]];

      // 3. تحويل الصفحة إلى JSON
      // defval: 0 تعني أن الخلايا الفارغة ستكون قيمتها 0 بدلاً من undefined
      const json = XLSX.utils.sheet_to_json(sheet, { defval: 0 });

      console.log("📊 البيانات الخام (JSON):", json);

      // إرجاع النتيجة
      resolve(json);
    };

    // معالجة الأخطاء (إضافة مهمة للأمان)
    reader.onerror = (error) => reject(error);

    // بدء عملية القراءة كـ ArrayBuffer
    reader.readAsArrayBuffer(file);
  });
}

// ==========================================
// دالة إنشاء وتنزيل ملف Excel الجديد
// ==========================================
function downloadExcel(rows) {
  // 1. تحويل مصفوفة البيانات (JSON) إلى ورقة عمل (Worksheet)
  const ws = XLSX.utils.json_to_sheet(rows);

  // 2. إنشاء كتاب عمل جديد (Workbook)
  const wb = XLSX.utils.book_new();

  // 3. إضافة الورقة إلى الكتاب وتسميتها "Phenix"
  XLSX.utils.book_append_sheet(wb, ws, "Phenix");

  // 4. حفظ الملف باسم "Phenix.xlsx" وتنزيله للمستخدم
  XLSX.writeFile(wb, "Phenix.xlsx");
}