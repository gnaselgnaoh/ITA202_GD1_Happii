# TÀI LIỆU GIẢI THÍCH MÃ NGUỒN – GIAI ĐOẠN 1 (`GiaiDoan1.ipynb`)
**Đề tài:** Dự báo giá đóng cửa cổ phiếu nhóm ngành Ngân hàng trên sàn HOSE  
**Sinh viên:** Hoàng Lê Minh Sang – **MSSV:** PK04346  

---

## Tổng quan luồng xử lý dữ liệu (Pipeline)
Toàn bộ mã nguồn trong `notebooks/GiaiDoan1.ipynb` được chia thành **5 Cell (5 bước)** nối tiếp nhau theo đúng quy trình chuẩn của một dự án Học máy (Machine Learning) trên dữ liệu chuỗi thời gian:
1. **Cell 1:** Đọc và hợp nhất 5 tệp Excel lịch sử giá từ CafeF $\rightarrow$ xuất `raw_data_nganhang.csv`.
2. **Cell 2:** Kiểm tra cấu trúc, phát hiện dữ liệu bẩn/bất thường và thống kê mô tả.
3. **Cell 3:** Làm sạch chuỗi thời gian, quy đổi giá điều chỉnh và tạo đặc trưng kỹ thuật (Feature Engineering) $\rightarrow$ xuất `clean_data_nganhang.csv`.
4. **Cell 4:** Trực quan hóa dữ liệu khám phá (EDA) và xuất 4 biểu đồ độ phân giải cao.
5. **Cell 5:** Mã hóa One-Hot, chia tập Train/Validation/Test theo mốc thời gian và chuẩn hóa đặc trưng (Scaling) chống rò rỉ dữ liệu.

---

## Giải thích chi tiết từng Cell code

### 1. Cell 1: Đọc & Hợp nhất dữ liệu thô (`Lab 1`)
* **Mục đích:** Gom dữ liệu của 5 ngân hàng (`ACB`, `BID`, `CTG`, `MBB`, `VCB`) từ nhiều file rời rạc về một bảng duy nhất để dễ quản lý và xử lý đồng loạt.
* **Giải thích các hàm chính:**
  * `glob.glob('../data/raw/CafeF_LichSuGia_HOSE_*.xlsx')`: Tự động quét và lấy đường dẫn của tất cả các file Excel khớp với mẫu tên trong thư mục `data/raw/`. Hàm `sorted()` đảm bảo danh sách file luôn được xếp theo thứ tự bảng chữ cái cố định.
  * `pd.read_excel(f, sheet_name='LichSuGia')`: Đọc đúng sheet chứa bảng lịch sử giao dịch của từng file.
  * `pd.concat(dfs, ignore_index=True)`: Nối chồng 5 bảng dữ liệu theo chiều dọc (1.230 dòng) và đánh lại chỉ số dòng liên tục từ `0` đến `1229`.
  * `raw_df.to_csv(..., encoding='utf-8-sig')`: Xuất ra file `.csv` thô. Tham số `utf-8-sig` giúp Microsoft Excel mở trực tiếp file tiếng Việt có dấu mà không bị lỗi font.

---

### 2. Cell 2: Kiểm tra chất lượng & Thống kê mô tả (`Lab 1 & Lab 2`)
* **Mục đích:** Khám phá tổng quan cấu trúc bảng, kiểm tra các lỗi dữ liệu thường gặp và trích xuất số liệu thống kê đưa vào báo cáo.
* **Giải thích các hàm chính:**
  * `raw_df.info()`: Hiển thị tên 12 cột, số lượng giá trị không rỗng (`Non-Null`) và kiểu dữ liệu (`Dtype`). Qua đó phát hiện cột `Ngày` đang ở dạng chuỗi (`str`) cần chuyển đổi.
  * `raw_df.isnull().sum()` & `raw_df.duplicated(subset=['Mã', 'Ngày']).sum()`: Kiểm tra giá trị khuyết thiếu và bản ghi trùng lặp theo khóa chính là cặp `(Mã, Ngày)`. Kết quả đều bằng `0`.
  * `(raw_df['Giá đóng cửa'] != raw_df['Giá điều chỉnh']).sum()`: Phát hiện **832 dòng** bị lệch giữa giá đóng cửa thực tế và giá điều chỉnh do chia cổ tức. Đây là cơ sở để viết logic quy đổi giá ở Cell 3.
  * `raw_df.groupby('Mã')[...].agg(...)`: Gom nhóm theo từng mã ngân hàng để tính riêng các chỉ số trung bình (`mean`), trung vị (`median`), nhỏ nhất (`min`), lớn nhất (`max`) và độ lệch chuẩn (`std`).

---

### 3. Cell 3: Làm sạch & Tạo đặc trưng mới (`Lab 2`)
* **Mục đích:** Xử lý sự cố đứt gãy giá do chia cổ tức, tạo các chỉ báo kỹ thuật phản ánh xu hướng thị trường và định nghĩa biến mục tiêu $y$.
* **Giải thích các hàm chính:**
  * `pd.to_datetime(df['Ngày'], format='%d/%m/%Y')`: Chuyển cột `Ngày` từ chuỗi `ngày/tháng/năm` sang kiểu thời gian chuẩn.
  * `df.sort_values(['Mã', 'Ngày'])`: **Cực kỳ quan trọng** — Dữ liệu tải từ CafeF mặc định xếp từ ngày mới nhất về cũ nhất. Bắt buộc phải sắp xếp tăng dần (cũ $\rightarrow$ mới) theo từng `Mã` trước khi tính toán các hàm trễ (`shift`) hay trung bình trượt (`rolling`).
  * `ty_le_dieu_chinh = df['Giá điều chỉnh'] / df['Giá đóng cửa']`: Tính hệ số điều chỉnh cổ tức tại từng phiên, sau đó nhân vào `Giá mở cửa`, `Giá cao nhất`, `Giá thấp nhất` để đồng bộ toàn bộ hệ quy chiếu giá (`Adj_Open`, `Adj_High`, `Adj_Low`, `Adj_Close`).
  * **Tạo 7 đặc trưng mới dùng `groupby('Mã')`:** Việc bọc `groupby('Mã')` đảm bảo khi tính toán cho mã này (ví dụ `BID`) không bị lấy nhầm dữ liệu của dòng cuối cùng thuộc mã trước đó (`ACB`).
    * `Price_Range`: Biên độ giá trong ngày (`Adj_High - Adj_Low`).
    * `Return = .pct_change()`: Tỷ suất sinh lời ngày so với phiên liền trước.
    * `MA5`, `MA20 = .rolling(5).mean()`, `.rolling(20).mean()`: Đường trung bình trượt 5 phiên và 20 phiên.
    * `Close_Lag1`, `Close_Lag2`, `Volume_Lag1 = .shift(1)`, `.shift(2)`: Kéo dữ liệu của 1 và 2 phiên quá khứ xuống dòng hiện tại làm đặc trưng đầu vào.
  * `df['Target_Close'] = df.groupby('Mã')['Adj_Close'].shift(-1)`: Dịch ngược giá đóng cửa của ngày mai (`t+1`) lên dòng hôm nay (`t`) để làm nhãn dự đoán ($y$).
  * `df.dropna()`: Loại bỏ 19 dòng đầu tiên của mỗi mã (do chưa đủ 20 ngày tính `MA20`) và 1 dòng cuối cùng của mỗi mã (do không có giá ngày mai cho `Target_Close`), thu được bảng sạch **1.130 dòng**.

---

### 4. Cell 4: Trực quan hóa dữ liệu khám phá — EDA (`Lab 3`)
* **Mục đích:** Vẽ 4 loại biểu đồ khác nhau để tìm ra quy luật ẩn trong dữ liệu và lưu thành file ảnh chất lượng cao (`dpi=300`) phục vụ báo cáo.
* **Giải thích từng biểu đồ:**
  * **Biểu đồ 1 (Line Plot – `1_xu_huong_gia.png`):** Vẽ đường diễn biến giá theo thời gian của 5 mã trên cùng một trục tọa độ để quan sát tính đồng pha và sự phân tầng mức giá.
  * **Biểu đồ 2 (Boxplot – `2_boxplot_return.png`):** Kiểm tra phân phối tỷ suất sinh lời ngày và phát hiện các điểm ngoại lai (outliers) chạm trần/sàn $\pm 7\%$ trên sàn HOSE.
  * **Biểu đồ 3 (Histogram + KDE – `3_phan_phoi_target.png`):** Khảo sát hình dạng phân phối của biến mục tiêu `Target_Close`, phát hiện cấu trúc 3 đỉnh (multimodal) quanh mốc 20k, 35k và 60k.
  * **Biểu đồ 4 (Correlation Heatmap – `4_heatmap_tuong_quan.png`):** Tính ma trận hệ số tương quan Pearson (`df_clean[cols_corr].corr()`) giữa các biến số, qua đó phát hiện hiện tượng đa cộng tuyến mạnh ($r = 0.99 - 1.00$) giữa nhóm biến giá.

---

### 5. Cell 5: Mã hóa, Chia tập & Chuẩn hóa dữ liệu (`Lab 4`)
* **Mục đích:** Biến đổi toàn bộ dữ liệu sang dạng ma trận số thuần túy, sẵn sàng đưa vào huấn luyện các thuật toán Học máy ở Giai đoạn 2.
* **Giải thích các hàm chính:**
  * `pd.get_dummies(..., columns=['Mã'], prefix='Ma', dtype=int)`: Thực hiện **One-Hot Encoding**, tách cột `Mã` thành 5 cột nhị phân (`0` hoặc `1`) để mô hình phân biệt được từng ngân hàng mà không gán thứ tự lớn/bé giả tạo.
  * **Chia tập theo mốc thời gian (70% – 15% – 15%):** Thay vì dùng `train_test_split` ngẫu nhiên (gây rò rỉ dữ liệu tương lai vào quá khứ), code lấy danh sách các ngày giao dịch duy nhất (`unique_dates`), cắt tại mốc 70% ngày đầu làm **Train** (790 dòng), 15% ngày tiếp theo làm **Validation** (170 dòng) và 15% ngày cuối cùng làm **Test** (170 dòng).
  * **Chuẩn hóa dữ liệu (`StandardScaler`):**
    * Chỉ chọn 12 cột biến số liên tục (`num_cols`) để chuẩn hóa về trung bình $0$, độ lệch chuẩn $1$ (giữ nguyên 5 cột nhị phân One-Hot).
    * `scaler.fit_transform(X_train[num_cols])`: Chỉ tính toán tham số trung bình và độ lệch chuẩn trên tập **Train**.
    * `scaler.transform(X_val[num_cols])` & `scaler.transform(X_test[num_cols])`: Dùng chính tham số của tập Train để biến đổi tập Val và Test, đảm bảo tuân thủ tuyệt đối nguyên tắc **chống rò rỉ dữ liệu (Anti-Data Leakage)**.