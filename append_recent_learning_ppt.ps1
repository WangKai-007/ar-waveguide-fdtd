$ErrorActionPreference='Stop'
$src='D:\FDTD_File\AR_Waveguide_RCWA_Phase_Learning_Summary_20260813.pptx'
$out='D:\FDTD_File\AR_Waveguide_RCWA_Recent_Learning_Summary_20260816.pptx'
Copy-Item -LiteralPath $src -Destination $out -Force
$pp=New-Object -ComObject PowerPoint.Application
$p=$pp.Presentations.Open($out)
$dark= [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb(17,34,56))
$blue= [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb(43,121,209))
$gray= [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb(96,112,128))
function txt($s,$t,$l,$top,$w,$h,$size,$color,$bold=$false){$x=$s.Shapes.AddTextbox(1,$l,$top,$w,$h);$x.TextFrame.TextRange.Text=$t;$x.TextFrame.TextRange.Font.Name='Aptos';$x.TextFrame.TextRange.Font.NameFarEast='Microsoft YaHei';$x.TextFrame.TextRange.Font.Size=$size;$x.TextFrame.TextRange.Font.Color.RGB=$color;$x.TextFrame.TextRange.Font.Bold=$(if($bold){-1}else{0});return $null}
function slide($title){$s=$p.Slides.Add($p.Slides.Count+1,12);$s.Background.Fill.ForeColor.RGB=16777215;txt $s $title 55 35 1150 45 30 $dark $true | Out-Null;return $s}
$s=slide '有限 9-ridge 扩瞳器：从场图走向截面功率';txt $s '完成内容' 70 115 300 30 24 $blue $true;txt $s "• 复制 9 个 ridge，建立 45%→70% 占空比渐变`n• 将 FDTD 区域扩展到覆盖整个阵列`n• 添加 gradient_field_monitor 查看 |E|² 与 Py`n• 添加 power_y_01…power_y_10 的 Y-normal 监视器" 85 165 1050 190 21 $dark;txt $s '关键认识：Z-normal 场监视器适合看场分布；Y-normal 功率监视器才适合比较截面功率。' 85 420 1030 55 21 $gray $true
$s=slide '0.8 µm 与 1.2 µm 间距：当前仍需抑制干涉';txt $s '0.8 µm 版本的 T（原始值）' 70 115 500 30 22 $blue $true;txt $s "T01=-0.4579  T02=-0.4807  T03=-0.4393  T04=-0.4724  T05=-0.5431`nT06=-0.4577  T07=0.5279  T08=0.5252  T09=0.5326  T10=0.5416" 80 165 1080 80 19 $dark;txt $s '1.2 µm 版本的 T（原始值）' 70 285 500 30 22 $blue $true;txt $s "T02=-0.9734  T03=-1.0033  T04=-0.9889  T05=-0.9882`nT06=0.0207  T07=0.0230  T08=0.0209  T09=0.0228  T10=0.0224" 80 335 1080 80 19 $dark;txt $s '结论：间距改变了功率分布，但 T 仍出现符号翻转与明显不连续；下一步应分离 T_forward/T_backward 或直接积分净功率。' 80 485 1080 60 20 $gray $true
$s=slide '当前判断与下一步验证';txt $s '当前判断' 70 115 280 30 24 $blue $true;txt $s "• 10 个 surface_normal=2，监视器方向一致`n• 负 extraction 不能解释为负效率，而是反射/干涉导致的净功率回流`n• 当前渐变设计尚未证明均匀出耦" 85 165 1050 130 21 $dark;txt $s '下一步' 70 360 280 30 24 $blue $true;txt $s "1. 读取每个监视器的 T_forward/T_backward（若结果可用）`n2. 对 Y-normal 的 power 在 x-z 面积分，而不是用 Z-normal 的局部 Py`n3. 比较 0.8/1.2 µm 间距的功率包络与均匀性`n4. 再优化 ridge 间距与占空比渐变" 85 410 1050 150 21 $dark
$p.Save();$p.Close();$pp.Quit();Write-Output $out
