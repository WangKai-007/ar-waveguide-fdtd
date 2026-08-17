$ErrorActionPreference = 'Stop'
trap { Write-Output $_.InvocationInfo.PositionMessage; Write-Output $_.Exception.Message; break }
Add-Type -AssemblyName System.Drawing
$out = 'D:\FDTD_File\AR_Waveguide_RCWA_Phase_Learning_Summary_20260813.pptx'
$render = 'D:\FDTD_File\ppt_rcwa_phase\render'
New-Item -ItemType Directory -Force -Path $render | Out-Null

$pp = New-Object -ComObject PowerPoint.Application
$pres = $pp.Presentations.Add()
$pres.PageSetup.SlideSize = 15 # 16:9 widescreen

$dark = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb(17,34,56))
$blue = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb(43,121,209))
$teal = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb(0,153,153))
$light = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb(245,248,252))
$gray = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb(96,112,128))
$line = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb(205,215,225))
$white = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::White)
$scale = 0.5625
$fontScale = 0.75

function Add-Text($s,$t,$l,$top,$w,$h,$size,$color,$bold=$false,$align=1) {
  $x = $s.Shapes.AddTextbox(1,$l*$scale,$top*$scale,$w*$scale,$h*$scale)
  $x.TextFrame.TextRange.Text = $t
  $x.TextFrame.TextRange.Font.NameFarEast = 'Microsoft YaHei'
  $x.TextFrame.TextRange.Font.Name = 'Aptos'
  $x.TextFrame.TextRange.Font.Size = $size*$fontScale
  $x.TextFrame.TextRange.Font.Color.RGB = $color
  $x.TextFrame.TextRange.Font.Bold = $(if($bold){-1}else{0})
  $x.TextFrame.TextRange.ParagraphFormat.Alignment = $align
  $x.TextFrame.MarginLeft=0; $x.TextFrame.MarginRight=0
  return $x
}
function Add-Rule($s,$x1,$y1,$x2,$y2,$color=$line,$weight=1.25) {
  $r=$s.Shapes.AddLine($x1*$scale,$y1*$scale,$x2*$scale,$y2*$scale);$r.Line.ForeColor.RGB=[int]$color;$r.Line.Weight=[single]($weight*$scale);return $r
}
function Add-Rect($s,$l,$top,$w,$h,$fill,$radius=$false) {
  $type=$(if($radius){5}else{1});$r=$s.Shapes.AddShape($type,$l*$scale,$top*$scale,$w*$scale,$h*$scale);$r.Fill.ForeColor.RGB=$fill;$r.Line.Visible=0;return $r
}
function New-Slide($title,$n) {
  $s=$pres.Slides.Add($pres.Slides.Count+1,12);$s.FollowMasterBackground=0;$s.Background.Fill.ForeColor.RGB=$white
  Add-Text $s 'AR WAVEGUIDE • RCWA PHASE' 45 24 330 18 11 $blue $true | Out-Null
  Add-Text $s $title 45 52 1130 48 30 $dark $true | Out-Null
  Add-Rule $s 45 112 1235 112 $line 0.9 | Out-Null
  Add-Text $s ('2026.08.13   |   '+$n) 1050 685 180 14 9 $gray $false 3 | Out-Null
  return $s
}

# 1 Cover
$s=$pres.Slides.Add(1,12);$s.FollowMasterBackground=0;$s.Background.Fill.ForeColor.RGB=$dark
Add-Text $s '阶段性学习总结' 65 64 500 26 16 $white $true | Out-Null
Add-Text $s 'AR 波导光栅：RCWA 快速优化与 RGB 色散' 65 160 900 118 42 $white $true | Out-Null
Add-Text $s '从 FDTD 基准验证，走到参数扫描、容差判断和三色周期反推' 67 304 840 40 21 $white $false | Out-Null
Add-Rect $s 68 407 80 7 $teal | Out-Null
Add-Text $s '学习阶段：RCWA 建模与工程化验证' 68 438 470 28 18 $white $false | Out-Null
Add-Text $s '2026.08.13' 68 642 250 22 14 $white $false | Out-Null

# 2 Workflow
$s=New-Slide '为什么在 FDTD 之后引入 RCWA？' '02'
Add-Text $s '核心目的：先快速筛选参数，再用 FDTD 做高可信验证。' 45 135 700 32 21 $gray $false | Out-Null
$steps=@(@('1','FDTD 基准','明确结构、场分布与能量守恒'),@('2','RCWA 对齐','确认级次效率与 FDTD 一致'),@('3','RCWA 扫描','快速找到占空比、高度、角度与波长趋势'),@('4','FDTD 复核','对候选点验证真实电磁场与制造影响'))
$x=55; foreach($q in $steps){Add-Rect $s $x 230 265 250 $light $true | Out-Null;Add-Text $s $q[0] ($x+18) 250 40 42 28 $blue $true | Out-Null;Add-Text $s $q[1] ($x+18) 310 210 35 24 $dark $true | Out-Null;Add-Text $s $q[2] ($x+18) 365 220 70 16 $gray $false | Out-Null;if($x -lt 900){Add-Rule $s ($x+266) 355 ($x+290) 355 $blue 2 | Out-Null};$x+=295}

# 3 validation
$s=New-Slide 'RCWA 与 FDTD 在 532 nm 工作点一致' '03'
Add-Text $s '基准结构：周期 386 nm，占空比 65%，高度 160 nm，TE 偏振，入射角 15°。' 45 135 1060 27 18 $gray | Out-Null
Add-Text $s '透射级次效率（相对入射光）' 60 205 420 24 19 $dark $true | Out-Null
$orders=@(@('-1 级','13.51%','13.66%'),@('0 级','71.43%','70.93%'),@('+1 级（目标）','9.80%','9.85%'))
Add-Text $s '级次' 60 250 180 22 16 $gray $true | Out-Null;Add-Text $s 'RCWA' 250 250 130 22 16 $gray $true | Out-Null;Add-Text $s 'FDTD' 405 250 130 22 16 $gray $true | Out-Null
$y=292;foreach($r in $orders){Add-Rule $s 60 ($y-8) 535 ($y-8) $line 0.8|Out-Null;Add-Text $s $r[0] 60 $y 180 23 17 $dark ($r[0] -like '+1*')|Out-Null;Add-Text $s $r[1] 250 $y 130 23 17 $blue ($r[0] -like '+1*')|Out-Null;Add-Text $s $r[2] 405 $y 130 23 17 $teal ($r[0] -like '+1*')|Out-Null;$y+=55}
Add-Rect $s 650 215 500 260 $light $true | Out-Null
Add-Text $s '验证结论' 685 248 180 28 24 $dark $true | Out-Null
Add-Text $s '+1 级仅差 0.045 个百分点。\nRCWA 分层、界面、偏振和周期设置可信。\n后续可用 RCWA 执行快速扫描。' 685 310 400 110 19 $dark $false | Out-Null

# 4 optimization
$s=New-Slide '粗扫与细扫都指向 65% / 160 nm 的局部最优点' '04'
Add-Text $s '占空比粗扫：45%–80%（步长 5%）；高度粗扫：80–240 nm（步长 20 nm）。' 45 135 1100 27 18 $gray | Out-Null
Add-Text $s '占空比扫描（高度固定 160 nm）' 60 200 420 24 19 $dark $true | Out-Null
$duty=@(45,50,55,60,65,70,75,80);$eta=@(4.74,5.87,7.88,9.37,9.80,9.43,8.52,7.11);$x=70;for($i=0;$i -lt $duty.Count;$i++){ $h=$eta[$i]*20;Add-Rect $s $x (470-$h) 38 $h $(if($duty[$i] -eq 65){$blue}else{$line}) |Out-Null;Add-Text $s ('{0}%' -f $duty[$i]) ($x-5) 485 55 18 12 $gray $false 2|Out-Null;$x+=55}
Add-Text $s '65%：9.80%' 255 270 170 22 18 $blue $true |Out-Null
Add-Text $s '高度细扫（占空比固定 65%）' 655 200 420 24 19 $dark $true |Out-Null
$height=@(140,150,160,170,180);$eh=@(9.48,9.74,9.80,9.67,9.37);$x=675;for($i=0;$i -lt $height.Count;$i++){ $h=$eh[$i]*20;Add-Rect $s $x (470-$h) 55 $h $(if($height[$i] -eq 160){$teal}else{$line}) |Out-Null;Add-Text $s ('{0} nm' -f $height[$i]) ($x-10) 485 75 18 12 $gray $false 2|Out-Null;$x+=80}
Add-Text $s '160 nm：9.80%' 835 270 190 22 18 $teal $true |Out-Null
Add-Text $s '容差观察：在 65% 下，150 / 170 nm 仍为 9.74% / 9.67%，对约 ±10 nm 深度偏差并不陡峭。' 60 580 1080 35 17 $gray |Out-Null

# 5 angle
$s=New-Slide '入射角改变时，+1 级会出现截止与效率下降' '05'
Add-Text $s '固定：532 nm、65% / 160 nm、TE 偏振。读数按级次 n=+1 自动匹配。' 45 135 1060 27 18 $gray|Out-Null
Add-Text $s '角度（°）' 80 555 90 18 14 $gray|Out-Null;Add-Text $s '+1 级效率' 88 205 140 22 18 $dark $true|Out-Null
Add-Rule $s 120 520 1080 520 $line 1|Out-Null;Add-Rule $s 120 230 120 520 $line 1|Out-Null
$ang=@(5,10,15,20,25);$ae=@(12.50,11.51,9.80,6.92,0);$prev=$null;for($i=0;$i -lt 5;$i++){ $x=190+$i*190;$y=520-$ae[$i]*20;Add-Rect $s ($x-6) ($y-6) 12 12 $blue $true|Out-Null;if($prev){Add-Rule $s $prev[0] $prev[1] $x $y $blue 2.2|Out-Null};$prev=@($x,$y);Add-Text $s ($ang[$i].ToString()) ($x-12) 535 25 18 14 $gray $false 2|Out-Null;Add-Text $s ('{0}%' -f $ae[$i].ToString('0.00')) ($x-28) ($y-36) 58 18 13 $dark $true 2|Out-Null}
Add-Rect $s 770 250 350 135 $light $true|Out-Null;Add-Text $s '工程含义' 800 275 200 25 22 $dark $true|Out-Null;Add-Text $s '效率不能脱离系统入射角单独优化。25° 时 +1 级截止；实际设计角需由耦入与波导几何共同确定。' 800 315 285 58 16 $gray|Out-Null

# 6 spectral
$s=New-Slide '单周期光栅对波长高度敏感，红光在 620 nm 截止' '06'
Add-Text $s '固定：θ=15°、65% / 160 nm、TE 偏振。每次运行强制锁定角度与波长。' 45 135 1100 27 18 $gray|Out-Null
$wl=@(460,490,520,532,550,580,620);$we=@(22.35,16.79,11.59,9.80,7.28,3.40,0);$x=100;for($i=0;$i -lt 7;$i++){ $h=$we[$i]*12;Add-Rect $s $x (510-$h) 70 $h $(if($wl[$i] -eq 532){$teal}else{$blue})|Out-Null;Add-Text $s ('{0} nm' -f $wl[$i]) ($x-5) 530 80 18 12 $gray $false 2|Out-Null;Add-Text $s ('{0}%' -f $we[$i].ToString('0.0')) ($x+5) (490-$h) 60 18 12 $dark $true 2|Out-Null;$x+=145}
Add-Text $s '蓝光强、绿光可用、红光截止：这是色偏与彩虹纹风险的直接数值证据。' 70 610 1050 28 20 $dark $true|Out-Null

# 7 RGB
$s=New-Slide 'RGB 要在同一波导角传播，必须使用不同周期' '07'
Add-Text $s '以绿光 532 nm 的 +1 级传播角 65.44° 为目标角：' 45 135 930 27 18 $gray|Out-Null
Add-Text $s 'sin θ+1 = [sin θin + λ / Λ] / ncore' 80 205 760 42 28 $dark $true|Out-Null
Add-Text $s 'ncore = 1.8；θin = 15°；临界角 = 33.75°；绿光 +1 级角 = 65.44°。' 80 260 900 25 17 $gray|Out-Null
$rgb=@(@('蓝 460 nm','333.75 nm',$blue),@('绿 532 nm','385.99 nm',$teal),@('红 620 nm','449.84 nm',16711680));$x=80;foreach($r in $rgb){Add-Rect $s $x 355 310 145 $light $true|Out-Null;Add-Rect $s $x 355 10 145 $r[2]|Out-Null;Add-Text $s $r[0] ($x+32) 385 220 24 20 $dark $true|Out-Null;Add-Text $s $r[1] ($x+32) 425 220 32 28 $r[2] $true|Out-Null;$x+=360}
Add-Text $s '结论：一个 386 nm 单周期光栅只对绿光匹配；RGB AR 波导通常需要多周期、多区域或多层设计。' 80 565 1070 36 19 $dark $true|Out-Null

# 8 closing
$s=New-Slide '阶段成果与下一步计划' '08'
Add-Text $s '已经建立：从物理模型到工程判断的完整学习闭环。' 45 135 950 32 21 $gray|Out-Null
$left=@('完成 RCWA 建模与分层界面设置','完成 RCWA–FDTD +1 级一致性验证','完成占空比、高度、角度与波长扫描','完成 RGB 周期反推计算器');$y=220;foreach($t in $left){Add-Rect $s 70 $y 13 13 $teal $true|Out-Null;Add-Text $s $t 105 ($y-5) 480 25 18 $dark|Out-Null;$y+=60}
Add-Rect $s 675 205 455 285 $light $true|Out-Null;Add-Text $s '下一步' 710 235 180 28 24 $dark $true|Out-Null;Add-Text $s '1. 为蓝/绿/红分别建立周期模型\n2. 扫描效率与角度容差\n3. 引入有限光栅与制造误差\n4. 用 FDTD 检查场分布与能量守恒\n5. 连接系统级 AR 光学设计' 710 285 350 150 18 $gray|Out-Null
Add-Text $s '当前保存的 RCWA 基准：65% duty / 160 nm height / 532 nm / θ=15° / TE。' 70 595 980 28 17 $gray|Out-Null

$pres.SaveAs($out)
for($i=1;$i -le $pres.Slides.Count;$i++){ $pres.Slides.Item($i).Export((Join-Path $render ('slide-{0}.png' -f $i)),'PNG',1280,720) }
$pres.Close();$pp.Quit()
Write-Output $out
