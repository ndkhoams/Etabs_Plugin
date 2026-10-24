# AGENTS.md — Quy tắc làm việc với repo Etabs_Plugin

Tài liệu cho các agent/người cùng phát triển plugin **Etabs Ultimate Tools**. Đọc hết trước khi sửa code.

## Giao tiếp & quyền ghi

- Giao tiếp với chủ repo (owner) bằng **tiếng Việt**, ngắn gọn, kèm bằng chứng kiểm chứng cho mỗi kết quả (commit, trạng thái build...).
- Owner đã cấp **pre-approval cho repo này** (từ 2026-10-23): agent được push trực tiếp, không cần hỏi duyệt trước từng lần. Điều kiện đi kèm: commit chỉ chứa đúng thay đổi thuộc phạm vi công việc đã thống nhất, và sau mỗi lần push phải báo lại cho owner kèm bằng chứng kiểm chứng (commit/blob SHA, trạng thái build). Các thao tác ngoài push thông thường (xóa branch, đổi settings repo, ghi đè lịch sử...) vẫn phải hỏi trước.
- Khi sửa lỗi: xác nhận đúng hàm/hành vi từ dữ kiện thật (đọc source, dữ liệu ETABS) trước khi kết luận — không đoán. Sửa xong phải rà soát **cả lớp lỗi tương tự** ở các file khác (ví dụ lỗi parse số theo culture thường lặp ở nhiều nơi).

## Bố cục repo — không tự ý đổi

- Toàn bộ mã nguồn nằm trong thư mục **`Ultimtate Tools/`** (tên thư mục viết sai chính tả "Ultimtate" là lịch sử của repo — **không đổi tên**, solution/csproj và quy trình phân phối đang dựa vào nó).
- Thư mục gốc chỉ gồm: `Etabs_Tool.iso` (bản phân phối), `README.md`, `AGENTS.md`, `.gitignore`.
- **Không đổi tên**: class `cPlugin` và method `Main` (điểm vào ETABS nạp plugin), các public class/row type, tên 9 tab, tên sheet Excel và định dạng các file xuất ra (text lực cột...) — người dùng và phần mềm ngoài đang dựa vào chúng.
- Hằng `currentBuild` (dấu thời gian build) **không còn nằm trong source**: target `StampPluginBuildTimestamp` trong csproj tự sinh file `BuildStamp.g.cs` trong `obj/` (thư mục gitignore) mỗi lần build — không sửa tay file sinh này, và build không còn làm bẩn cây làm việc (không còn chuyện phải Undo trước khi pull).
- **Không đổi cơ chế cập nhật**: tab Check for Updates lấy commit mới nhất của `Etabs_Tool.iso` trên GitHub (GitHub API), tải ISO tại đúng commit đó và so bản build đang chạy với **`version.txt` đóng gói ở thư mục gốc bên trong ISO** (ISO cũ chưa có version.txt thì lùi về so theo ngày commit của file ISO). Phát hành bản mới = build → đóng gói ISO mới — bên trong **bắt buộc có `version.txt`** ở thư mục gốc: file này được csproj **tự xuất vào thư mục output cạnh DLL sau mỗi lần build** (target `WritePluginVersionFile`, chứa đúng dấu build `yyyyMMdd-HHmmss` của lần build đó), khi đóng gói chỉ cần nhặt kèm file có sẵn trong `bin/<cấu hình>/` — đè lên `Etabs_Tool.iso` ở root → push. Đóng gói thiếu version.txt là lỗi quy trình phát hành.

## Build & chuẩn "xong"

- Project SDK-style, **.NET Framework 4.8**, WinForms, x64; solution: `Ultimtate Tools/Etabs_Ultimate_Tools.slnx`. Tham chiếu `ETABSv1.dll` theo đường dẫn cài ETABS 22 mặc định; thư viện ClosedXML 0.102.2 qua NuGet.
- Chuẩn hoàn thành một thay đổi: **build sạch (0 lỗi) trong Visual Studio trên máy có ETABS 22 + chạy thử tính năng trong ETABS thật**. Máy Linux của agent không có ETABS nên không thay được bước chạy thử trong ETABS — phải nói rõ giới hạn này khi bàn giao.
- Agent **có thể build kiểm chứng biên dịch** trên Linux (đã làm được ngày 2026-10-23) theo đúng công thức sau, và chỉ trên **bản sao** của project — tuyệt đối không sửa csproj gốc để phục vụ việc này:
  1. Cài .NET SDK 8 (gói tarball chính thức, giải nén bằng `tar --no-same-owner`; script `dotnet-install.sh` từng lỗi trên máy này).
  2. Lấy file `ETABSv1.dll` từ owner (DLL trong thư mục cài ETABS 22 của owner) — không tự chế/không dùng DLL lạ.
  3. Trên bản sao: trỏ `HintPath` của reference ETABSv1 sang DLL đó, thêm `PackageReference Microsoft.NETFramework.ReferenceAssemblies 1.0.3` (PrivateAssets=all). (Từ 2026-10-24, target `StampPluginBuildTimestamp` đã dùng task MSBuild thuần và chạy được trên Linux, không còn cần gỡ.)
  4. `dotnet build` cả Debug và Release phải đạt 0 lỗi / 0 cảnh báo trước khi push.

## Quy tắc đúng số liệu (bài học từ audit 2026-10-23 — bắt buộc)

Đây là công cụ kiểm tra kết cấu; kết quả sai mà không báo là lỗi nghiêm trọng nhất. Các quy tắc sau đã được rút ra từ các lỗi thật từng tồn tại trong code:

1. **Parse số từ UI luôn dùng `EtabsHelper.TryParseNumber`** — không bao giờ dùng `double.TryParse` mặc định (theo CurrentCulture): trên Windows tiếng Việt, `"3.12"` bị parse thành `312`. Helper này hiểu cả `3.12`, `3,12`, `1.234,5` và `1.000` = 1000. Giá trị hệ số không hợp lệ hoặc ≤ 0 phải **cảnh báo và dừng**, không lặng lẽ thay bằng 1.0.
2. **Thiếu dữ liệu thì không kết luận đạt**: khi tầng/tổ hợp không có drift, lực cắt, khối lượng, phản lực... phải ghi **"Không có dữ liệu"** / **"Không xác định"**, không dùng giá trị 0 để tính tiếp ra kết luận "OK"/"Đạt". Trạng thái này phải nhất quán từ extractor → lưới hiển thị → file Excel xuất ra.
3. **Khôi phục trạng thái ETABS sau mỗi lần chạy**: mọi tính năng chạm vào trạng thái toàn cục của phiên ETABS (đơn vị hiện hành, lựa chọn case/combo của Results.Setup, `SetOptionMultiValuedCombo`, lựa chọn hiển thị Database Tables) phải được bọc trong `using (new EtabsStateGuard(_sap)) { ... }` ở điểm vào của form (mẫu: các `Run*` trong `ModelCheckForm.Checks.cs`, preview ở ColumnExport và PileH).
4. **Kiểm tra return code của API ETABS** trước khi dùng kết quả, đặc biệt các hàm trả mảng song song (`Results.StoryDrifts`, `Results.JointReact`, `Results.PierForce`...): `ret != 0` hoặc mảng null/thiếu phần tử = không có dữ liệu, không index mù.
5. **Ghi file xuất luôn dùng `CultureInfo.InvariantCulture`** cho số (file text dùng dấu phẩy tách cột — format theo culture máy sẽ phá cột) và **UTF-8** cho nội dung có tên tiếng Việt (không dùng ASCII).
6. **Lực/phản lực đi theo đúng dòng dữ liệu của nó**: không trộn max của thành phần này ở step/dòng này với max của thành phần khác ở step/dòng khác rồi kết luận. Với tổ hợp envelope không đọc được step-by-step, giá trị suy ra chỉ là cận dưới → đánh dấu "Không xác định", không khẳng định "bảo thủ" khi chưa chứng minh được.
7. Tên pier/case/combo do ETABS trả về là dữ kiện thật: **không cắt ngắn/ép khuôn** chúng khi làm khóa tra cứu (từng có lỗi cắt tên pier tại dấu `-` khiến `V-1` và `V-2` bị gộp thành một pier). Chỉ chuẩn hóa tối thiểu (trim, hoa/thường, hậu tố `@story` nếu có).

## Chữ ký API ETABS 22 đã kiểm chứng (tránh nhầm với SAP2000)

- `cSapModel.GetPresentUnits()` — **không tham số**, trả `eUnits` trực tiếp (không phải dạng `ref` như SAP2000).
- `Results.Setup.SetCaseSelectedForOutput(name, bool)` / `SetComboSelectedForOutput(name, bool)` — 2 tham số; các getter tương ứng là `(name, ref bool)`.
- `Results.Setup.GetOptionMultiValuedCombo(ref int)` / `SetOptionMultiValuedCombo(int)`.
- `DatabaseTables.Get/SetLoadCasesSelectedForDisplay` và `...Combinations...` theo dạng `(ref int, ref string[])` / `(ref string[])`.
- Khi nghi ngờ một chữ ký: đọc metadata của chính `ETABSv1.dll` bản ETABS 22 (MetadataLoadContext) thay vì suy từ tài liệu SAP2000/ETABS bản khác.
