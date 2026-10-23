# Etabs Ultimate Tools

Plugin cho **ETABS 22** hỗ trợ kiểm tra kết cấu theo tiêu chuẩn Việt Nam: chuyển vị, chuyển vị lệch tầng (gió & động đất), hiệu ứng P-Delta, lực dọc cột/vách, phản lực cọc và xuất nội lực — tất cả trong một form duy nhất, đọc trực tiếp kết quả phân tích từ model ETABS đang mở.

![ETABS](https://img.shields.io/badge/ETABS-22-blue) ![.NET Framework](https://img.shields.io/badge/.NET%20Framework-4.8-purple) ![Platform](https://img.shields.io/badge/Platform-Windows%20x64-lightgrey)

## Tính năng

Form chính gồm 9 tab:

| Tab | Chức năng |
|---|---|
| **Property Modifiers** | Gán hệ số giảm độ cứng cho các frame/area đang chọn trong ETABS, có chụp lại trạng thái gốc để khôi phục (rollback). |
| **Displacements** | Chuyển vị đỉnh theo tổ hợp gió, đối chiếu giới hạn H/500. Nguồn số liệu: điểm tầng (Story) hoặc tâm khối lượng sàn (Diaphragm), chọn cách gộp theo mode lớn nhất/trung bình. |
| **Wind Drifts** | Chuyển vị lệch tầng do gió, giới hạn mặc định 1/500. Đọc từ bảng Story Drifts hoặc suy ra từ chuyển vị tầng. |
| **Seismic Drifts** | Chuyển vị lệch tầng do động đất theo **TCVN 9386-1:2025**: điều kiện `dr · ν ≤ limit · h`, với limit chọn 0.005 / 0.0075 / 0.010, hệ số q và ν nhập tay. |
| **P-Delta** | Hệ số ổn định θ = P·Δ / (V·h) theo từng tầng: P từ bảng Mass Summary, V từ Story Forces, Δ từ Story Drifts; kết luận theo các ngưỡng θ ≤ 0.10 / 0.20 / 0.30, có hệ số khuếch đại 1/(1−θ). |
| **Axial Force** | Kiểm tra tỉ số nén νd = NEd / (Ac · fcd) cho cột (frame) và vách (pier) đang chọn trong ETABS. Vách được phân loại Pier.Wall / Pier.Col theo tỉ lệ cạnh tiết diện; giới hạn mặc định 0.65 (cột) / 0.40 (vách), fcd suy từ cấp bền bê tông nhập trên form (αcc = 1.0, γc = 1.2). |
| **Column Force Exporter** | Xuất nội lực (PU, MUXT, MUYT, MUXB, MUYB) của cột/vách đang chọn theo nhiều tổ hợp cùng lúc, ra Excel hoặc file text (định dạng nạp cho CSI Col / Prokon). |
| **Pile Reactions** | Kiểm tra khả năng chịu tải của cọc (kéo / nén / ngang H) cho 3 trường hợp tải: Tải đứng, Tải gió, Động đất. Sức chịu tải nhập theo từng loại cọc (point spring); lực ngang lấy theo step có H lớn nhất khi ETABS cho phép đọc step-by-step. |
| **Check for Updates** | Kiểm tra bản build mới trên GitHub và tải file ISO bản phân phối mới nhất về máy. |

Hầu hết các tab đều xuất được báo cáo **Excel (.xlsx)** để lưu hồ sơ tính toán.

### Hành vi khi chạy kiểm tra

- Plugin tự chuyển đơn vị ETABS về **kN–m** khi tính và **khôi phục lại đơn vị, lựa chọn case/combo xuất kết quả và các tùy chọn hiển thị của ETABS về đúng trạng thái trước khi chạy** khi tính xong.
- Tầng/tổ hợp nào **thiếu dữ liệu** (chưa Run Analysis, không có bảng kết quả...) sẽ được ghi rõ **"Không có dữ liệu"** / **"Không xác định"** thay vì kết luận đạt — hãy kiểm tra lại model trong các trường hợp này.
- Mọi số nhập trên form (q, ν, sức chịu tải cọc...) đều đọc được cả kiểu `3.12`, `3,12` lẫn `1.234,5` bất kể ngôn ngữ Windows.

## Yêu cầu

- **ETABS 22** (bản cài đặt chính thức của CSI) trên Windows x64.
- .NET Framework 4.8 (có sẵn trên Windows 10/11 hiện hành).

## Cài đặt & sử dụng

Bản phân phối được đóng gói trong file **`Etabs_Tool.iso`** ở thư mục gốc của repo:

1. Tải/mở file `Etabs_Tool.iso` (Windows 10/11 mount trực tiếp bằng double-click) và chép thư mục plugin ra vị trí cố định trên máy.
2. Trong ETABS: **Tools → Add/Show Plugins… → Browse…** trỏ tới file `Etabs_Ultimate_Tools.dll`, đặt tên hiển thị tùy ý, nhấn **Add** rồi **Done**.
3. Mở model, **Run Analysis** trước khi dùng các tab kiểm tra (riêng các tab Axial Force / Property Modifiers / Column Force Exporter thì **chọn trước** cột/vách/point cọc cần xử lý trong ETABS).
4. Chạy plugin từ menu **Tools** của ETABS, chọn tab tương ứng, chọn tổ hợp tải và bấm nút Tính/Xem trước.

## Cập nhật

Mở tab **Check for Updates** trong plugin: công cụ tải file `Etabs_Tool.iso` ở commit mới nhất trên GitHub và so sánh bản build đang chạy với dấu build ghi trong **`version.txt` đóng gói bên trong chính file ISO đó** — một nguồn sự thật duy nhất, không lệch giữa code và bản phân phối. Có bản mới thì nút tải mới bật; file ISO vừa tải về để kiểm tra được dùng lại luôn cho bước tải, không phải tải hai lần. (Tải về không tự cài — thay thư mục plugin bằng bản trong ISO mới như bước cài đặt.)

## Build từ mã nguồn

1. Cần **Visual Studio 2022** và **ETABS 22 đã cài** trên máy (project tham chiếu `ETABSv1.dll` theo đường dẫn cài đặt mặc định `C:\Program Files\Computers and Structures\ETABS 22\`).
2. Mở solution `Ultimtate Tools/Etabs_Ultimate_Tools.slnx`, build (F6) cấu hình Debug hoặc Release. DLL kết quả nằm trong `Ultimtate Tools/bin/<cấu hình>/`.
3. Project là SDK-style, target `net48`, thư viện xuất báo cáo dùng ClosedXML 0.102.2 (NuGet tự restore).
4. Mỗi lần build, csproj **tự đóng dấu thời gian build** vào hằng `currentBuild` trong `ModelCheckForm.Update.cs` (định dạng `yyyyMMdd-HHmmss`) — đây là "phiên bản" hiển thị trên tiêu đề form và dùng cho tab cập nhật. Không sửa tay hằng này. Đồng thời csproj tự xuất file **`version.txt`** (chứa đúng dấu build đó) vào thư mục output cạnh DLL — khi đóng gói ISO, nhặt kèm file này vào thư mục gốc của ISO (xem mục Cập nhật).

## Cấu trúc thư mục

```
Etabs_Plugin/
├── Etabs_Tool.iso          # Bản phân phối đóng gói để cài đặt/cập nhật
└── Ultimtate Tools/        # Mã nguồn (tên thư mục giữ nguyên như lịch sử repo)
    ├── Etabs_Ultimate_Tools.slnx / .csproj
    ├── cPlugin.cs          # Điểm vào plugin (ETABS gọi class cPlugin, method Main)
    ├── ModelCheckForm*.cs  # Form chính và các partial theo từng tab
    ├── *Extractor.cs       # Đọc/khớp kết quả ETABS (drift, chuyển vị, P-Delta...)
    ├── *Exporter.cs        # Xuất báo cáo Excel / file text
    ├── Etabshelper.cs      # Tiện ích dùng chung + EtabsStateGuard (lưu/khôi phục trạng thái ETABS)
    └── EtabsTableReader.cs # Đọc bảng kết quả Database Tables của ETABS
```

## Tác giả

KhoaND13 — © theo dấu build hiển thị trên tiêu đề form.
