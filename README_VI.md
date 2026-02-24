# SheetBrain 🧠 - Hướng dẫn Tiếng Việt

**SheetBrain** là công cụ phân tích và tự động hóa Excel thông minh được hỗ trợ bởi AI. Công cụ cung cấp kiến trúc 3 giai đoạn (Hiểu-Thực thi-Xác thực) để phân tích dữ liệu Excel toàn diện với khả năng cải thiện lặp lại.

> 🚀 **Mới sử dụng SheetBrain?** Xem [Hướng dẫn nhanh](QUICKSTART.md) để bắt đầu trong 5 phút!

## Tính năng

- 🔍 **Module Hiểu biết**: Phân tích cấu trúc và ngữ cảnh Excel sử dụng LLM đa phương thức
- 💻 **Module Thực thi**: Lý luận đa vòng và thực thi code Python cho phân tích phức tạp
- ✅ **Module Xác thực**: Đảm bảo chất lượng và cải thiện lặp lại thông qua xác thực LLM
- 📊 **Bộ công cụ Excel**: Tiện ích toàn diện cho các thao tác Excel (đọc, ghi, định dạng, biểu đồ)
- 🔄 **Cải thiện Lặp lại**: Tự động tinh chỉnh phân tích thông qua phản hồi xác thực
- 🌐 **Hỗ trợ Đa ngôn ngữ**: Phản hồi bằng Tiếng Việt, Tiếng Anh và nhiều ngôn ngữ khác
- 🛠️ **Cấu hình Linh hoạt**: Cài đặt và tùy chọn triển khai có thể tùy chỉnh

## Cài đặt Nhanh

### 1. Clone và Cài đặt

```bash
git clone https://github.com/microsoft/SheetBrain.git
cd SheetBrain
pip install -e .
```

### 2. Thiết lập Google Cloud

```bash
# Tạo service account (thay your-project-id bằng project ID của bạn)
make create-sa PROJECT_ID=your-project-id
```

### 3. Cấu hình

```bash
# Tạo file .env
make env

# Chỉnh sửa .env với project ID của bạn
nano .env
```

Cập nhật trong `.env`:
```bash
GOOGLE_PROJECT_ID=your-actual-project-id
LANGUAGE=Vietnamese  # Để nhận phản hồi bằng Tiếng Việt
```

## Sử dụng Nhanh

### Ví dụ 1: Phân tích với Tiếng Việt

```bash
# Sử dụng command line
python main.py data.xlsx "Tổng doanh thu là bao nhiêu?" --language Vietnamese

# Hoặc chạy demo tiếng Việt
python demo_vietnamese.py

# Hoặc sau khi cài đặt
sheetbrain-demo-vi
```

### Ví dụ 2: Sử dụng Python Code

```python
from core.agent import SheetBrain
from config.settings import Config

# Cấu hình với tiếng Việt
config = Config(
    service_account_path="service-account.json",
    language="Vietnamese"  # Phản hồi bằng Tiếng Việt
)

# Khởi tạo agent
agent = SheetBrain(excel_path="dulieu.xlsx", config=config)

# Đặt câu hỏi bằng tiếng Việt
result = agent.run("Tìm 5 khách hàng có doanh thu cao nhất")

# Kết quả
print(f"Câu trả lời: {result['answer']}")
print(f"Độ tin cậy: {result['confidence_score']:.2f}")
print(f"Thành công: {result['success']}")
```

### Ví dụ 3: Các loại câu hỏi

```bash
# Truy vấn đơn giản
python main.py dulieu.xlsx "Tổng doanh thu là bao nhiêu?" --language Vietnamese

# Phân tích
python main.py dulieu.xlsx "Top 5 sản phẩm bán chạy nhất?" --language Vietnamese

# So sánh
python main.py dulieu.xlsx "So sánh doanh thu Q1 và Q2" --language Vietnamese

# Phức tạp
python main.py dulieu.xlsx "Phân tích xu hướng bán hàng theo tháng" \
    --language Vietnamese \
    --max-turns 5
```

## Các Câu hỏi Mẫu

### Về Doanh thu
- "Tổng doanh thu trong năm 2023 là bao nhiêu?"
- "Doanh thu trung bình mỗi tháng?"
- "Tháng nào có doanh thu cao nhất?"

### Về Sản phẩm
- "Sản phẩm nào bán chạy nhất?"
- "Top 10 sản phẩm theo doanh số?"
- "Sản phẩm nào có lợi nhuận cao nhất?"

### Về Khách hàng
- "Tổng số khách hàng?"
- "Khách hàng nào mua nhiều nhất?"
- "Phân tích hành vi mua hàng của khách hàng"

### Về Xu hướng
- "Xu hướng tăng trưởng theo quý?"
- "So sánh năm 2022 và 2023"
- "Dự đoán doanh thu quý tiếp theo"

## Hiểu Kết quả

Khi chạy phân tích, bạn sẽ thấy:

```
🚀 [SheetBrain] Bắt đầu phân tích 3 giai đoạn lặp lại...
================================================================================
📖 [GIAI ĐOẠN 1] MODULE HIỂU BIẾT
----------------------------------------
✅ [GIAI ĐOẠN 1] Hoàn thành trong 5.23s

🔄 [VÒNG LẶP 1/3] CHU TRÌNH THỰC THI-XÁC THỰC
============================================================
💻 [VÒNG LẶP 1] MODULE THỰC THI
----------------------------------------
✅ [VÒNG LẶP 1] Thực thi hoàn thành trong 7.45s

🔍 [VÒNG LẶP 1] MODULE XÁC THỰC
----------------------------------------
✅ [VÒNG LẶP 1] Xác thực hoàn thành trong 2.11s
🎯 [VÒNG LẶP 1] Độ tin cậy: 0.95

============================================================
KẾT QUẢ PHÂN TÍCH
============================================================
Thành công: ✅
Câu trả lời: Tổng doanh thu Q4 2023 là 1.234.567.890 đồng
Độ tin cậy: 0.95/1.0
Số vòng lặp: 1
Thời gian: 14.79s
============================================================
```

## Cấu hình Nâng cao

### Sử dụng file .env

```bash
# .env
GOOGLE_SERVICE_ACCOUNT_PATH=service-account.json
GOOGLE_PROJECT_ID=your-project-id
GOOGLE_LOCATION=us-central1
GOOGLE_MODEL_NAME=gemini-2.0-flash-exp
LANGUAGE=Vietnamese
MAX_TURNS=5
TOKEN_BUDGET=10000
```

### Sử dụng Python Config

```python
from config.settings import Config

config = Config(
    service_account_path="service-account.json",
    project_id="my-project",
    location="us-central1",
    model_name="gemini-2.0-flash-exp",
    language="Vietnamese",
    max_turns=5,
    enable_validation=True,
    enable_understanding=True
)
```

## Các Ngôn ngữ Được Hỗ trợ

- **English** (Tiếng Anh) - Mặc định
- **Vietnamese** (Tiếng Việt) 🇻🇳
- **Chinese** (中文) 🇨🇳
- **Japanese** (日本語) 🇯🇵
- **Korean** (한국어) 🇰🇷
- **Spanish** (Español) 🇪🇸
- **French** (Français) 🇫🇷
- **German** (Deutsch) 🇩🇪

## Xử lý Lỗi

### Lỗi xác thực
```
Lỗi: Could not determine credentials
```
**Giải pháp:**
1. Kiểm tra file `service-account.json` có tồn tại
2. Đặt biến môi trường: `export GOOGLE_APPLICATION_CREDENTIALS="$(pwd)/service-account.json"`

### Chưa có Project ID
```
Lỗi: Project ID is required
```
**Giải pháp:**
```bash
export GOOGLE_PROJECT_ID="your-project-id"
```

### API chưa kích hoạt
```
Lỗi: Vertex AI API has not been used
```
**Giải pháp:**
```bash
gcloud services enable aiplatform.googleapis.com
```

## Tài liệu

- **[README.md](README.md)** - Tài liệu đầy đủ (Tiếng Anh)
- **[QUICKSTART.md](QUICKSTART.md)** - Hướng dẫn nhanh
- **[INSTALL.md](INSTALL.md)** - Hướng dẫn cài đặt chi tiết
- **[MIGRATION.md](MIGRATION.md)** - Hướng dẫn chuyển đổi từ OpenAI

## Ví dụ Nâng cao

### Phân tích với hình ảnh

```python
from PIL import Image

# Tải ảnh chụp màn hình Excel
image = Image.open("excel_screenshot.png")

result = agent.run(
    user_question="Phân tích biểu đồ trong hình này",
    table_image=image
)
```

### Phân tích lặp

```python
# Agent tự động lặp lại khi xác thực thất bại
result = agent.run(
    user_question="Phân tích phức tạp cần nhiều bước",
    max_turns=10,  # Cho phép nhiều vòng lặp hơn
    enable_validation=True  # Kích hoạt cải thiện tự động
)

print(f"Tổng số vòng lặp: {result['total_iterations']}")
```

## Đóng góp

Dự án này hoan nghênh các đóng góp và đề xuất. Xem [CONTRIBUTING.md](CONTRIBUTING.md) để biết thêm chi tiết.

## Trích dẫn

Nếu bạn thấy công việc của chúng tôi hữu ích, vui lòng trích dẫn:

```bibtex
@article{wang2025sheetbrain,
  title   = {SheetBrain: A Neuro-Symbolic Agent for Accurate Reasoning over Complex and Large Spreadsheets},
  author  = {Wang, Ziwei and Su, Jiayuan and Zhou, Mengyu and Zeng, Huaxing and Jia, Mengni and Lv, Xiao and Dong, Haoyu and Ma, Xiaojun and Han, Shi and Zhang, Dongmei},
  journal = {arXiv preprint arXiv:2510.19247},
  year    = {2025}
}
```

## Giấy phép

Dự án này được cấp phép theo Giấy phép MIT - xem file [LICENSE](LICENSE) để biết chi tiết.

## Hỗ trợ

- 📧 Email: opencode@microsoft.com
- 🐛 Báo lỗi: [GitHub Issues](https://github.com/microsoft/SheetBrain/issues)
- 📖 Tài liệu: [README.md](README.md)

---

Được phát triển với ❤️ bởi Microsoft Research
