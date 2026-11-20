$ErrorActionPreference = 'Stop'
$template = 'C:\Users\PC\.codex\plugins\cache\openai-curated-remote\openai-templates\0.1.1\skills\artifact-template-project-kickoff\assets\reference.pptx'
$root = (Get-Location).Path
$outputDir = Join-Path $root 'reports'
$buildDir = Join-Path $root '.codex-build'
New-Item -ItemType Directory -Force -Path $outputDir,$buildDir | Out-Null
$out = Join-Path $outputDir 'Assignment_GD1_ThuyetTrinh.pptx'
if (Test-Path -LiteralPath $out) { $out = Join-Path $outputDir 'Assignment_GD1_ThuyetTrinh_20261120.pptx' }
$chart = Join-Path $root 'reports\figures\1_xu_huong_gia.png'
if (!(Test-Path -LiteralPath $template)) { throw 'Không tìm thấy template Project Kickoff.' }
if (!(Test-Path -LiteralPath $chart)) { throw 'Không tìm thấy biểu đồ xu hướng giá.' }
function Set-TextContains($slide, [string]$needle, [string]$value, [int]$occurrence = 1) {
  $matches = @()
  for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
    $shape = $slide.Shapes.Item($i)
    try { if ($shape.HasTextFrame -and $shape.TextFrame.HasText) { if ([string]$shape.TextFrame.TextRange.Text -like "*$needle*") { $matches += $shape } } } catch {}
  }
  if ($matches.Count -lt $occurrence) { return $false }
  $matches[$occurrence - 1].TextFrame.TextRange.Text = $value
  return $true
}
function Add-TextBox($slide, [string]$text, [double]$left, [double]$top, [double]$width, [double]$height, [double]$size, [int]$color) {
  $box = $slide.Shapes.AddTextbox(1, $left, $top, $width, $height)
  $box.TextFrame.MarginLeft = 0; $box.TextFrame.MarginRight = 0; $box.TextFrame.MarginTop = 0; $box.TextFrame.MarginBottom = 0
  $box.TextFrame.WordWrap = -1
  $box.TextFrame.TextRange.Text = $text
  $box.TextFrame.TextRange.Font.Size = $size
  $box.TextFrame.TextRange.Font.Color.RGB = $color
  return $box
}
$dark = 0x263A2E
$ppt = $null; $pres = $null
try {
  $ppt = New-Object -ComObject PowerPoint.Application
  $ppt.Visible = 0; $ppt.DisplayAlerts = 0
  $pres = $ppt.Presentations.Open($template, -1, -1, 0)
  $pres.SaveAs($out, 24)
  while ($pres.Slides.Count -gt 7) { $pres.Slides.Item($pres.Slides.Count).Delete() }
  $nl = [string][char]10
  $s = $pres.Slides.Item(1)
  [void](Set-TextContains $s 'Team Project' ("DỰ BÁO GIÁ CỔ PHIẾU" + $nl + "NGÀNH NGÂN HÀNG"))
  [void](Set-TextContains $s 'Aligning on the vision' 'Giá đóng cửa điều chỉnh phiên kế tiếp · HOSE')
  [void](Set-TextContains $s '2026 H2 Planning' 'ITA202 · MÁY HỌC CƠ BẢN')
  Add-TextBox $s 'Hoàng Lê Minh Sang  ·  PK04346' 52 472 760 30 15 $dark | Out-Null

  $s = $pres.Slides.Item(2)
  [void](Set-TextContains $s 'Today’s Agenda' 'Bài toán: dự đoán phiên kế tiếp')
  $body = @('ĐẦU VÀO','Giá và khối lượng giao dịch của một mã đến phiên t','','ĐẦU RA','Giá đóng cửa điều chỉnh của phiên giao dịch kế tiếp t+1','','PHẠM VI','ACB · BID · CTG · MBB · VCB trên HOSE','','DẠNG BÀI TOÁN','Học có giám sát · hồi quy chuỗi thời gian · dự báo một phiên') -join $nl
  $b = Add-TextBox $s $body 58 132 830 390 23 $dark
  $b.TextFrame.TextRange.ParagraphFormat.SpaceAfter = 8

  $s = $pres.Slides.Item(3)
  [void](Set-TextContains $s 'Half text, half image slide' 'Dữ liệu và phạm vi')
  [void](Set-TextContains $s 'Title here' ("Nguồn & quy mô" + $nl + "CafeF · 5 mã · 5 tệp Excel" + $nl + "1.230 dòng · 12 cột") 1)
  [void](Set-TextContains $s 'Title here' ("Chất lượng dữ liệu" + $nl + "0 giá trị thiếu · 0 khóa trùng" + $nl + "832 dòng lệch giá đóng cửa/giá điều chỉnh") 2)
  for ($i = 1; $i -le $s.Shapes.Count; $i++) { $shape = $s.Shapes.Item($i); if ($shape.Type -eq 13) { $shape.Delete(); break } }
  $img = $s.Shapes.AddPicture($chart, 0, -1, 7.0*72, 2.0*72, 5.9*72, 2.46*72)
  $img.AlternativeText = 'Biểu đồ xu hướng giá điều chỉnh của ACB, BID, CTG, MBB và VCB'

  $s = $pres.Slides.Item(4)
  [void](Set-TextContains $s 'Large message slides' 'Tiền xử lý và tạo đặc trưng')
  $callouts = @(
    ("Điều chỉnh giá" + $nl + "Quy đổi giá mở, cao, thấp theo giá điều chỉnh để xử lý ảnh hưởng cổ tức."),
    ("Đặc trưng lịch sử" + $nl + "Return, MA5, MA20, Close_Lag1/2, Volume_Lag1 và Price_Range."),
    ("Nhãn dự đoán" + $nl + "Target_Close là Adj_Close của phiên t+1. Tập sạch còn 1.130 dòng.")
  )
  for ($i = 1; $i -le 3; $i++) { [void](Set-TextContains $s 'Title here' $callouts[$i-1] $i) }
  $steps = @('BƯỚC 1','BƯỚC 2','BƯỚC 3')
  for ($i = 1; $i -le 3; $i++) { [void](Set-TextContains $s 'Supertitle' $steps[$i-1] $i) }

  $s = $pres.Slides.Item(5)
  [void](Set-TextContains $s 'Four points grid layout' 'Khám phá dữ liệu: bốn quan sát')
  $insights = @(
    ("Mặt bằng giá" + $nl + "Năm mã nằm ở các vùng giá khác nhau; cần nhận diện riêng từng mã."),
    ("Lợi suất ngày" + $nl + "Phần lớn quan sát gần 0%, cùng một số phiên biến động mạnh."),
    ("Tương quan" + $nl + "Giá hiện tại, giá trễ và MA có tương quan cao với giá phiên sau."),
    ("Hàm ý mô hình" + $nl + "Tương quan giữa biến giá cao; nên thử Ridge bên cạnh mô hình tuyến tính.")
  )
  for ($i = 1; $i -le 4; $i++) { [void](Set-TextContains $s 'Title goes here' $insights[$i-1] $i) }

  $s = $pres.Slides.Item(6)
  [void](Set-TextContains $s 'One column layout with checklist' 'Chia tập và chuẩn hóa dữ liệu')
  [void](Set-TextContains $s 'Lorem ipsum dolor sit amet, consectetur adipiscing elit ' 'One-hot encode mã cổ phiếu để mô hình phân biệt 5 ngân hàng.')
  [void](Set-TextContains $s 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Pellentesque ultricies quam ut massa luctus cursus.' 'Chuẩn hóa 12 biến số liên tục. Chỉ fit scaler trên Train để tránh rò rỉ thông tin tương lai.')
  [void](Set-TextContains $s 'Title goes here' ("Chia theo thời gian" + $nl + "Train 70% · Validation 15% · Test 15%") 1)
  [void](Set-TextContains $s 'Title goes here' ("Chống rò rỉ dữ liệu" + $nl + "Fit StandardScaler trên Train; dùng tham số đó cho Validation và Test.") 2)

  $s = $pres.Slides.Item(7)
  [void](Set-TextContains $s 'Intro with four topics' 'Kết luận và hướng tiếp theo')
  [void](Set-TextContains $s 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Pellentesque ultricies quam ut massa luctus cursus.' 'Giai đoạn 1 đã hoàn thành phần chuẩn bị dữ liệu. Notebook Giai đoạn 2 chưa có kết quả huấn luyện; các mô hình dưới đây là kế hoạch.')
  $topics = @(
    @('Thử mô hình','Linear Regression, Ridge, Random Forest và XGBoost.'),
    @('Đánh giá','MAE, RMSE, R²; so sánh với baseline giá phiên sau ≈ giá hiện tại.'),
    @('Phạm vi dự báo','Giá đóng cửa điều chỉnh của một phiên giao dịch kế tiếp.'),
    @('Cách sử dụng','Tham khảo phân tích ngắn hạn; không thay thế quyết định đầu tư.')
  )
  $topicShapes = @(); $descShapes = @()
  for ($i = 1; $i -le $s.Shapes.Count; $i++) {
    $shape = $s.Shapes.Item($i)
    try {
      if ($shape.HasTextFrame -and $shape.TextFrame.HasText -and [string]$shape.TextFrame.TextRange.Text -like '*Item or topic*') { $topicShapes += $shape }
      elseif ($shape.HasTextFrame -and $shape.TextFrame.HasText -and [string]$shape.TextFrame.TextRange.Text -like '*Lorem ipsum*' -and $shape.Left -gt 300) { $descShapes += $shape }
    } catch {}
  }
  $topicShapes = @($topicShapes | Sort-Object Top); $descShapes = @($descShapes | Sort-Object Top)
  for ($i=0; $i -lt 4; $i++) {
    if ($topicShapes.Count -gt $i) { $topicShapes[$i].TextFrame.TextRange.Text = $topics[$i][0] }
    if ($descShapes.Count -gt $i) { $descShapes[$i].TextFrame.TextRange.Text = $topics[$i][1] }
  }
  $pres.Save()
  for ($i=1; $i -le 7; $i++) { $pres.Slides.Item($i).Export((Join-Path $buildDir ("slide-{0}.png" -f $i)), 'PNG', 1600, 900) }
  Write-Output "PPTX=$out"; Write-Output "SLIDES=$($pres.Slides.Count)"; Write-Output "PREVIEWS=$buildDir"
}
finally { if ($pres -ne $null) { $pres.Close() }; if ($ppt -ne $null) { $ppt.Quit() } }
