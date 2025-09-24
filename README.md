# 🔍 ScriptHelper.RegExp – COM DLL for VBA Regular Expressions

**ScriptHelper.RegExp** là một COM DLL viết bằng Delphi, thay thế hoàn toàn **VBScript.RegExp**, đặc biệt hữu ích trên các hệ điều hành Windows mới (10/11) nơi Microsoft đã loại bỏ `vbscript.dll`.

Công cụ này giúp lập trình VBA (Excel, Word, Access) thực hiện các thao tác xử lý chuỗi bằng **Regular Expression (Regex)** một cách nhanh chóng và Unicode-friendly.

---

## 📌 Tính năng chính

* **Pattern**: định nghĩa biểu thức chính quy.
* **Global**: tìm tất cả kết quả trong chuỗi.
* **IgnoreCase**: bỏ qua phân biệt chữ hoa/thường.
* **Execute()**: trả về danh sách các match.
* **Replace()**: thay thế toàn bộ match.
* **FirstMatch()**: lấy match đầu tiên.
* **ReplaceFirst()**: thay thế match đầu tiên.
* **Split()**: tách chuỗi theo biểu thức chính quy.

---

## ⚡ Một vài ví dụ điển hình

### 1️⃣ Tìm số năm 4 chữ số

```vb
Dim re As Object, matches As Variant
Set re = CreateObject("ScriptHelper.RegExp")
re.Pattern = "\b\d{4}\b"
re.Global = True
matches = re.Execute("Year 2025, month 09, day 23")
Debug.Print matches(0)   ' Output: 2025
```

### 2️⃣ Thay thế tất cả chữ hoa bằng "X"

```vb
Dim re As Object
Set re = CreateObject("ScriptHelper.RegExp")
re.Pattern = "[A-Z]"
re.Global = True
Debug.Print re.Replace("AbCdeFG", "X")  ' Output: XbXdexXx
```

### 3️⃣ Tách chuỗi theo dấu phẩy hoặc khoảng trắng

```vb
Dim re As Object, parts As Variant
Set re = CreateObject("ScriptHelper.RegExp")
re.Pattern = "[, ]+"
parts = re.Split("A,B C,D")
Debug.Print parts(0)  ' Output: A
Debug.Print parts(1)  ' Output: B
```

---

## ✅ Ưu điểm

* Hoạt động ổn định trên **Windows 10/11** mà không cần VBScript.
* Hỗ trợ **Unicode** đầy đủ.
* Tương thích với cú pháp quen thuộc của **VBScript.RegExp**.
* Dùng được cả **Excel, Word, Access**.

---

## 🎯 Kết luận

Nếu bạn đang duy trì macro VBA cũ hoặc cần Regex trong VBA trên Windows mới, **ScriptHelper.RegExp** là giải pháp đơn giản, gọn nhẹ, dễ tích hợp và ổn định.

Các ví dụ chi tiết hơn đã có trong file **Demo VBA** đi kèm DLL, bạn có thể mở ra để thử tất cả tính năng: tìm kiếm, thay thế, tách chuỗi, kiểm tra ngày tháng, email, URL…


