#!/usr/bin/env bash

# File: publish.sh
# Hướng dẫn: Chạy bằng lệnh `bash publish.sh` để đẩy file lên GitHub và lấy link trực tuyến.

set -e

echo "=========================================================="
echo "🚀 ĐANG KHỞI CHẠY TIẾN TRÌNH TỰ ĐỘNG UP FILE LÊN GITHUB 🚀"
echo "=========================================================="

# 1. Kiểm tra Git đã được khởi tạo chưa
if [ ! -d .git ]; then
    echo "📂 Thư mục hiện tại chưa được cấu hình Git. Đang thiết lập..."
    git init
    echo "✅ Đã khởi tạo Git repository cục bộ."
fi

# 2. Đảm bảo nhánh chính tên là main
git branch -M main

# 3. Kiểm tra Remote Origin
REMOTE_URL=$(git remote get-url origin 2>/dev/null || echo "")
if [ -z "$REMOTE_URL" ]; then
    echo "🔗 Chưa liên kết với GitHub repository."
    echo "👉 Vui lòng nhập link HTTPS repository của bạn trên GitHub"
    echo "   (Ví dụ: https://github.com/username/task-seo.git):"
    read -r input_url
    if [ -z "$input_url" ]; then
        echo "❌ Lỗi: Bạn cần cung cấp URL repository để tiếp tục."
        exit 1
    fi
    git remote add origin "$input_url"
    REMOTE_URL="$input_url"
    echo "✅ Đã liên kết với Remote Origin: $REMOTE_URL"
fi

# 4. Kiểm tra Git User Config
USER_NAME=$(git config user.name || echo "")
USER_EMAIL=$(git config user.email || echo "")

if [ -z "$USER_NAME" ]; then
    echo "👤 Chưa thiết lập Tên người dùng Git (user.name)."
    echo "👉 Nhập tên hiển thị của bạn (Ví dụ: Nguyen Tuong Vy):"
    read -r input_name
    git config user.name "$input_name"
fi

if [ -z "$USER_EMAIL" ]; then
    echo "📧 Chưa thiết lập Email người dùng Git (user.email)."
    echo "👉 Nhập email của bạn (Ví dụ: email@domain.com):"
    read -r input_email
    git config user.email "$input_email"
fi

# 5. Kiểm tra trạng thái thay đổi
echo "🔍 Đang kiểm tra các file thay đổi..."
# Tạo một tệp .gitignore cơ bản nếu chưa có
if [ ! -f .gitignore ]; then
    echo ".DS_Store" > .gitignore
    echo "node_modules/" >> .gitignore
    echo "*.log" >> .gitignore
    echo "✅ Đã tạo tệp .gitignore mặc định."
fi

git add .

# Kiểm tra xem có gì để commit không
if git diff --cached --quiet; then
    echo "ℹ️ Không có thay đổi nào mới cần đẩy lên GitHub."
else
    echo "💾 Đang tiến hành commit các thay đổi..."
    git commit -m "Cập nhật nội dung báo cáo SEO [$(date '+%Y-%m-%d %H:%M:%S')]"
fi

# 6. Đẩy code lên GitHub
echo "📤 Đang đẩy dữ liệu lên GitHub (Có thể yêu cầu nhập Token nếu chưa đăng nhập)..."
git push -u origin main

# 7. Tính toán và hiển thị link trực tuyến
echo ""
if [[ "$REMOTE_URL" =~ github\.com[:/]([^/]+)/([^/.]+)(\.git)?$ ]]; then
    GH_USER="${BASH_REMATCH[1]}"
    GH_REPO="${BASH_REMATCH[2]}"
    
    # Chuẩn hóa tên repo nếu kết thúc bằng .git
    GH_REPO="${GH_REPO%.git}"
    
    PAGES_URL="https://${GH_USER}.github.io/${GH_REPO}/"
    
    echo "=========================================================="
    echo "🎉 ĐÃ UP LÊN GITHUB THÀNH CÔNG!"
    echo "🔗 Link trang chủ trực tuyến của bạn:"
    echo "👉 $PAGES_URL"
    echo "----------------------------------------------------------"
    echo "💡 Đường dẫn trực tiếp của từng file HTML:"
    
    # Tìm các file html trong thư mục (loại trừ các thư mục ẩn)
    find . -maxdepth 2 -name "*.html" -not -path '*/.*' | while read -r html_file; do
        clean_path="${html_file#./}"
        echo "🔗 $PAGES_URL$clean_path"
    done
    echo "=========================================================="
    echo "⚠️  Lưu ý: Nếu đây là lần đầu tiên bạn tạo Repo, hãy chắc chắn"
    echo "    đã bật GitHub Pages trong phần Settings -> Pages của Repo đó."
    echo "=========================================================="
else
    echo "⚠️  Không thể tự động tạo link GitHub Pages do định dạng URL không quen thuộc: $REMOTE_URL"
fi
