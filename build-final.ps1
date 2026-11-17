# Builds the final deck by injecting hand-written slide XML into a
# PowerPoint-generated valid package shell.
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$dir    = 'C:\Users\Acer\Documents\Event launch ai'
$shell  = "$env:TEMP\_shell.pptx"          # PowerPoint-generated 1-slide deck
$outFile = Join-Path $dir 'Event Launch RSVP - User Manual.pptx'

# ---- 1. fresh shell from PowerPoint itself ----
$pp = New-Object -ComObject PowerPoint.Application
$pres = $pp.Presentations.Add($false)
$pres.Slides.Add(1, 1) | Out-Null
$pres.Slides.Add(2, 1) | Out-Null
$pres.Slides.Add(3, 1) | Out-Null
$pres.Slides.Add(4, 1) | Out-Null
$pres.Slides.Add(5, 1) | Out-Null
$pres.Slides.Add(6, 1) | Out-Null
$pres.Slides.Add(7, 1) | Out-Null
$pres.Slides.Add(8, 1) | Out-Null
$pres.SaveAs($shell)
# save rels map: slide N -> rIdN via presentation part; PowerPoint orders sldIdLst = creation order
$pres.Close(); $pp.Quit()
Write-Output 'shell created'

# ---- 2. generate slide XML with a clean renderer ----
$script:SN = 100

function Esc([string]$s) { $s.Replace('&','&amp;').Replace('<','&lt;').Replace('>','&gt;') }

function Tx([double]$x,[double]$y,[double]$w,[double]$h,[double]$sz,[string]$color,[string]$text,[string]$align='l',[bool]$bold=$false,[bool]$italic=$false){
  $script:SN++
  $b = if($bold){"1"}else{"0"}; $i = if($italic){"1"}else{"0"}
  $anchor = if($align -eq 'ctr'){"ctr"}else{"t"}
  $t = Esc $text
  '<p:sp><p:nvSpPr><p:cNvPr id="'+$script:SN+'" name="tx"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="'+[int]($x*914400)+'" y="'+[int]($y*914400)+'"/><a:ext cx="'+[int]($w*914400)+'" cy="'+[int]($h*914400)+'"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr wrap="square" anchor="'+$anchor+'"><a:normAutofit/></a:bodyPr><a:lstStyle/><a:p><a:pPr algn="'+$align+'"/><a:r><a:rPr lang="en-US" sz="'+[int]($sz*100)+'" b="'+$b+'" i="'+$i+'" dirty="0"><a:solidFill><a:srgbClr val="'+$color+'"/></a:solidFill><a:latin typeface="Segoe UI"/></a:rPr><a:t>'+$t+'</a:t></a:r></a:p></p:txBody></p:sp>'
}

function Box([double]$x,[double]$y,[double]$w,[double]$h,[string]$fill,[string]$line,[double]$radius,[double]$lw=1.0){
  $script:SN++
  $f = if($fill){'<a:solidFill><a:srgbClr val="'+$fill+'"/></a:solidFill>'}else{'<a:noFill/>'}
  $l = if($line){'<a:ln w="'+[int]($lw*12700)+'"><a:solidFill><a:srgbClr val="'+$line+'"/></a:solidFill></a:ln>'}else{'<a:ln><a:noFill/></a:ln>'}
  $g = 'rect'; $av = '<a:avLst/>'
  if($radius -gt 0){ $g='roundRect'; $av='<a:avLst><a:gd name="adj" fmla="val '+[int]($radius*100000)+'"/></a:avLst>' }
  '<p:sp><p:nvSpPr><p:cNvPr id="'+$script:SN+'" name="sh"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="'+[int]($x*914400)+'" y="'+[int]($y*914400)+'"/><a:ext cx="'+[int]($w*914400)+'" cy="'+[int]($h*914400)+'"/></a:xfrm><a:prstGeom prst="'+$g+'">'+$av+'</a:prstGeom>'+$f+$l+'</p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>'
}

function Rule([double]$x,[double]$y,[double]$w,[string]$fill){
  $script:SN++
  '<p:sp><p:nvSpPr><p:cNvPr id="'+$script:SN+'" name="ln"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="'+[int]($x*914400)+'" y="'+[int]($y*914400)+'"/><a:ext cx="'+[int]($w*914400)+'" cy="0"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom><a:ln w="12700"><a:solidFill><a:srgbClr val="'+$fill+'"/></a:solidFill></a:ln></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>'
}

function Dot([double]$x,[double]$y,[double]$d,[string]$fill){
  $script:SN++
  '<p:sp><p:nvSpPr><p:cNvPr id="'+$script:SN+'" name="dot"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="'+[int]($x*914400)+'" y="'+[int]($y*914400)+'"/><a:ext cx="'+[int]($d*914400)+'" cy="'+[int]($d*914400)+'"/></a:xfrm><a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="'+$fill+'"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>'
}

function Chip([double]$x,[double]$y,[string]$n,[string]$fill){
  (Box $x $y 0.42 0.42 $fill '' 0.5) + (Tx $x ($y+0.045) 0.42 0.34 14 'FFFFFF' $n 'ctr' $true)
}

# palette
$DKBG='22423A'; $PRIM='3A6B5C'; $SOFT='E7F0EC'; $NO='A4553F'; $NOSF='F6E9E4'
$INK='2B2B2B'; $MUT='7A7A7A'; $LINE='E4E4E0'; $LTG='B9D3C9'; $MDG='8FBFAD'; $CARD='2C5147'

function Wrap([string]$body,[bool]$dark){
  $bg = if($dark){ $DKBG } else { 'FFFFFF' }
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="'+$bg+'"/></a:solidFill><a:effectLst/></p:bgPr></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="12192000" cy="6858000"/></a:xfrm></p:grpSpPr>'+$body+'</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>'
}

# ===== slide bodies =====
# S1 cover
$b = ''
$b += Box 0.9 0.9 1.6 0.52 $PRIM '' 0.5
$b += Tx 0.9 0.975 1.6 0.4 13 'FFFFFF' 'USER MANUAL' 'ctr' $true
$b += Tx 0.9 1.7 11.5 1.15 54 'FFFFFF' 'Event Launch RSVP' 'l' $true
$b += Tx 0.9 2.95 9.6 0.7 20 $LTG 'How to submit your reply, browse the guest list, and read the live headcount'
$cards = @(
  @('1','RSVP Form','Tell the host who you are and whether you are coming.'),
  @('2','Guest List','See every reply so far, newest first.'),
  @('3','Headcount Summary','Live totals confirmed by the replies.')
)
$cx = 0.9
foreach($c in $cards){
  $b += Box $cx 4.05 3.75 1.85 $CARD '' 0.06
  $b += Tx ($cx+0.3) 4.35 1.0 0.6 30 $MDG $c[0] 'l' $true
  $b += Tx ($cx+0.3) 5.0 3.15 0.4 16 'FFFFFF' $c[1] 'l' $true
  $b += Tx ($cx+0.3) 5.38 3.15 0.5 12 $LTG $c[2]
  $cx += 3.92
}
$b += Tx 0.9 6.6 8 0.35 12 $MDG 'Event Launch team  |  Version 1.0'
$slidesOut = @(Wrap $b $true)

# S2 what the page is
$b = ''
$b += Tx 0.7 0.55 8 0.6 30 $INK 'What this page is' 'l' $true
$b += Tx 0.7 1.2 11.9 0.6 15 $MUT 'One single web page (index.html) with three linked sections - every reply updates the totals instantly.'
$steps = @(
  @('1','RSVP Form','You fill it in and submit'),
  @('2','Guest List','Your reply appears here'),
  @('3','Headcount Summary','Totals recount automatically')
)
for($i=0;$i -lt 3;$i++){
  $y = 2.15 + $i*1.6
  $b += Chip 0.8 ($y+0.36) $steps[$i][0] $PRIM
  $b += Tx 1.45 ($y+0.28) 4.0 0.4 18 $INK $steps[$i][1] 'l' $true
  $b += Tx 1.45 ($y+0.66) 4.6 0.35 13 $MUT $steps[$i][2]
}
$b += Box 6.6 2.05 6.0 1.7 $SOFT '' 0.06
$b += Tx 7.0 2.35 5.2 0.4 17 $PRIM 'Guests' 'l' $true
$b += Tx 7.0 2.78 5.2 0.75 13 $INK 'Open the link the host sent you and reply in Section 1. No account or login is needed.'
$b += Box 6.6 3.95 6.0 1.7 $SOFT '' 0.06
$b += Tx 7.0 4.25 5.2 0.4 17 $PRIM 'The host' 'l' $true
$b += Tx 7.0 4.68 5.2 0.75 13 $INK 'Watches Sections 2 and 3 on the same page - replies and headcount update live as guests submit.'
$b += Tx 0.7 6.55 11.9 0.4 12 $MUT 'Note: replies live in the page itself; the sample entries you see on first open are placeholders.' 'l' $false $true
$slidesOut += Wrap $b $false

# S3 tour
$b = ''
$b += Tx 0.7 0.55 9 0.6 30 $INK 'A tour of the page, top to bottom' 'l' $true
$b += Tx 0.7 1.18 11.9 0.5 15 $MUT 'The layout mirrors how you will actually use it: reply first, then check, then read the numbers.'
$wx = 0.8; $ww = 5.4
$b += Box $wx 1.95 $ww 4.9 'FFFFFF' $LINE 0.05
$b += Tx ($wx+0.35) 2.15 4.5 0.35 14 $INK 'Event Launch RSVP' 'l' $true
$b += Tx ($wx+0.35) 2.5 4.7 0.3 9 $MUT 'Reply below - the host sees every answer and the live headcount on this page.'
$b += Box ($wx+0.35) 2.95 4.7 1.5 $SOFT '' 0.04
$b += Tx ($wx+0.55) 3.1 4.3 0.3 11 $PRIM '1  RSVP Form' 'l' $true
$b += Box ($wx+0.55) 3.5 4.3 0.28 'FFFFFF' $LINE 0.1
$b += Tx ($wx+0.75) 3.54 3.5 0.2 9 $MUT 'Your name'
$b += Box ($wx+0.55) 3.9 2.05 0.28 $PRIM '' 0.3
$b += Tx ($wx+0.55) 3.94 2.05 0.2 9 'FFFFFF' 'Yes, I''m coming' 'ctr' $true
$b += Box ($wx+2.8) 3.9 2.05 0.28 'FFFFFF' $LINE 0.3
$b += Tx ($wx+2.8) 3.94 2.05 0.2 9 $MUT 'No, can''t make it' 'ctr'
$b += Box ($wx+0.35) 4.75 4.7 0.8 'FFFFFF' $LINE 0.04
$b += Tx ($wx+0.55) 4.87 4.3 0.3 11 $PRIM '2  Guest List' 'l' $true
$b += Box ($wx+0.35) 5.75 4.7 0.85 $SOFT '' 0.04
$b += Tx ($wx+0.55) 5.88 4.3 0.3 11 $PRIM '3  Headcount Summary' 'l' $true
$tour = @(
  @('Header','Page title and one-line instructions.'),
  @('Section 1 - RSVP Form','Name, Yes/No toggle, pax, slot, days, invited by, and the Add my reply button.'),
  @('Section 2 - Guest List','Cards for each reply: a green YES or terracotta NO badge plus pax, slot, days and inviter.'),
  @('Section 3 - Headcount Summary','Three big stat tiles and a per-slot pax breakdown table.')
)
for($i=0;$i -lt 4;$i++){
  $y = 2.1 + $i*0.95
  $b += Dot 6.75 ($y+0.12) 0.16 $PRIM
  $b += Tx 7.05 $y 5.6 0.35 15 $INK $tour[$i][0] 'l' $true
  $b += Tx 7.05 ($y+0.32) 5.6 0.6 12.5 $MUT $tour[$i][1]
}
$slidesOut += Wrap $b $false

# S4 steps
$b = ''
$b += Tx 0.7 0.55 9 0.6 30 $INK 'Submitting your reply' 'l' $true
$b += Tx 0.7 1.18 11.9 0.5 15 $MUT 'Five steps in Section 1. Fields marked optional can be left blank.'
$st = @(
  @('Type your name','The only required field. Example placeholder: Nurul Hidayah.'),
  @('Pick Yes or No','Tap one of the two buttons - Yes turns green, No turns terracotta. Exactly one stays selected.'),
  @('Set total pax','How many people are coming including you (0-20). Only counts if you answered Yes.'),
  @('Choose slot and days','Morning / Afternoon / Evening slot, and 1-3 days. Both optional.'),
  @('Add my reply','Optionally note who invited you, then press the green button. Your card appears at the top of the Guest List.')
)
for($i=0;$i -lt 5;$i++){
  $y = 1.95 + $i*1.02
  $b += Chip 0.8 ($y+0.05) ([string]($i+1)) $PRIM
  $b += Tx 1.42 $y 4.4 0.4 16 $INK $st[$i][0] 'l' $true
  $b += Tx 1.42 ($y+0.36) 5.2 0.62 12.5 $MUT $st[$i][1]
  if($i -lt 4){ $b += Rule 1.01 ($y+0.62) 0.0 $LINE }
}
$b += Box 7.0 2.0 5.6 2.2 $NOSF '' 0.06
$b += Tx 7.35 2.25 5.0 0.4 16 $NO 'If the form complains' 'l' $true
$b += Tx 7.35 2.68 5.0 1.4 12.5 $INK 'A terracotta message appears - "Please fill in your name and pick yes or no" - when the name is empty or no Yes/No button was selected. Fill both and submit again.'
$b += Box 7.0 4.45 5.6 1.85 $SOFT '' 0.06
$b += Tx 7.35 4.7 5.0 0.4 16 $PRIM 'Answered No?' 'l' $true
$b += Tx 7.35 5.12 5.0 1.0 12.5 $INK 'Pax, slot and days are ignored - your guest card simply shows your name, a NO badge, and who invited you.'
$slidesOut += Wrap $b $false

# S5 guest list
$b = ''
$b += Tx 0.7 0.55 9 0.6 30 $INK 'Reading the Guest List' 'l' $true
$b += Tx 0.7 1.18 11.9 0.5 15 $MUT 'Section 2 shows every reply, newest first. Each card carries the same five details.'
$b += Box 0.8 1.95 5.7 1.55 'FFFFFF' $LINE 0.06
$b += Tx 1.1 2.15 3.0 0.35 15 $INK 'Farid Amran' 'l' $true
$b += Box 4.6 2.16 0.9 0.3 $SOFT '' 0.5
$b += Tx 4.6 2.19 0.9 0.24 11 $PRIM 'YES' 'ctr' $true
$b += Tx 1.1 2.72 5.2 0.55 11.5 $MUT 'Pax: 4      Slot: Morning (10am - 12pm)      Days: 2      Invited by: Aisyah Rahman'
$b += Box 0.8 3.7 5.7 1.55 'FFFFFF' $LINE 0.06
$b += Tx 1.1 3.9 3.2 0.35 15 $INK 'Siti Aminah Bakar' 'l' $true
$b += Box 4.6 3.91 0.9 0.3 $NOSF '' 0.5
$b += Tx 4.6 3.94 0.9 0.24 11 $NO 'NO' 'ctr' $true
$b += Tx 1.1 4.47 5.2 0.55 11.5 $MUT 'Pax: -      Slot: -      Days: -      Invited by: Daniel Cheang'
$notes = @(
  @('Green YES badge','A confirmed party - its pax feeds the summary totals.'),
  @('Terracotta NO badge','A polite decline - counted in No replies only.'),
  @('Newest first','Your own reply always lands at the top of the list.'),
  @('Dashes mean skipped','A - appears when a field was left blank or does not apply.')
)
for($i=0;$i -lt 4;$i++){
  $y = 2.0 + $i*0.95
  $b += Dot 7.0 ($y+0.1) 0.14 $PRIM
  $b += Tx 7.3 $y 5.3 0.35 15 $INK $notes[$i][0] 'l' $true
  $b += Tx 7.3 ($y+0.32) 5.3 0.6 12.5 $MUT $notes[$i][1]
}
$slidesOut += Wrap $b $false

# S6 summary
$b = ''
$b += Tx 0.7 0.55 9 0.6 30 $INK 'Reading the Headcount Summary' 'l' $true
$b += Tx 0.7 1.18 11.9 0.5 15 $MUT 'Section 3 recounts automatically every time anyone submits. No button to press.'
$tiles = @( @('0.8','6','Confirmed pax (yes replies)'), @('3.3','2','Yes replies (parties)'), @('5.8','1','No replies') )
foreach($t in $tiles){
  $tx = [double]$t[0]
  $b += Box $tx 1.95 2.3 1.5 $SOFT '' 0.06
  $b += Tx $tx 2.15 2.3 0.75 40 $PRIM $t[1] 'ctr' $true
  $b += Tx ($tx+0.1) 2.98 2.1 0.4 11 $MUT $t[2] 'ctr'
}
$b += Box 0.8 3.75 7.8 2.5 'FFFFFF' $LINE 0.04
$b += Tx 1.15 3.95 3.0 0.3 11 $MUT 'Slot' 'l' $true
$b += Tx 6.4 3.95 1.9 0.3 11 $MUT 'CONFIRMED PAX' 'r' $true
$slots = @( @('Morning (10am - 12pm)','4'), @('Afternoon (2pm - 4pm)','0'), @('Evening (7pm - 9pm)','2') )
for($i=0;$i -lt 3;$i++){
  $y = 4.35 + $i*0.58
  $b += Rule 1.15 $y 7.1 $LINE
  $b += Tx 1.15 ($y+0.08) 4.8 0.32 12.5 $INK $slots[$i][0]
  $b += Tx 6.9 ($y+0.08) 1.4 0.32 13 $INK $slots[$i][1] 'r' $true
}
$b += Tx 9.0 2.55 3.6 0.35 13 $INK 'Confirmed pax per slot' 'l' $true
$bars = @( @('Morning',4,0.0), @('Afternoon',0,1.05), @('Evening',2,2.1) )
foreach($bar in $bars){
  $by = 3.1 + [double]$bar[2]
  $b += Tx 9.0 $by 1.5 0.3 11 $INK $bar[0]
  $bh = [double]$bar[1] * 0.35
  if($bh -gt 0){ $b += Box 10.6 ($by+0.02) $bh 0.28 $PRIM '' 0.0 }
  $b += Tx 11.75 $by 0.5 0.3 11 $INK $bar[1]
}
$b += Tx 0.8 6.5 8.0 0.35 12 $MUT 'Sample figures shown; totals follow whatever the guest list contains.' 'l' $false $true
$slidesOut += Wrap $b $false

# S7 good to know
$b = ''
$b += Tx 0.7 0.55 9 0.6 30 $INK 'Good to know' 'l' $true
$b += Tx 0.7 1.18 11.9 0.5 15 $MUT 'Practical limits and small behaviours worth knowing before you use the page.'
$kn = @(
  @('No accounts, no logins','Anyone with the link can reply. Do not use it for private guest data.'),
  @('Pax is capped at 20','The number field accepts 0-20. Larger parties should reply as several entries.'),
  @('One answer per button pair','Choosing Yes deselects No automatically - you cannot pick both.'),
  @('Entries cannot be edited or deleted','There is no undo. Double-check pax and slot before pressing Add my reply.'),
  @('Sample data on first open','The page ships with three sample replies so the host can see the layout immediately.'),
  @('Works on phone and desktop','The layout is responsive - guests can reply from the event WhatsApp link.')
)
for($i=0;$i -lt 6;$i++){
  $x = @(0.8, 6.9)[[math]::Floor($i/3)]
  $y = 2.0 + ($i % 3)*1.55
  $b += Box $x $y 5.6 1.35 'FFFFFF' $LINE 0.06
  $b += Tx ($x+0.3) ($y+0.18) 5.0 0.35 14.5 $PRIM $kn[$i][0] 'l' $true
  $b += Tx ($x+0.3) ($y+0.56) 5.0 0.7 12 $INK $kn[$i][1]
}
$slidesOut += Wrap $b $false

# S8 closing
$b = ''
$b += Tx 0.9 1.5 11.5 1.0 44 'FFFFFF' 'You are ready to reply.' 'l' $true
$b += Tx 0.9 2.7 10 0.45 18 $LTG 'Quick checklist before you press Add my reply:'
$chk = @('Your name is filled in','Yes or No is selected','Pax, slot and days are correct (if coming)')
for($i=0;$i -lt 3;$i++){
  $y = 3.35 + $i*0.62
  $b += Dot 1.0 ($y+0.1) 0.2 $MDG
  $b += Tx 1.4 $y 9 0.4 16 'FFFFFF' $chk[$i]
}
$b += Box 0.9 5.5 5.9 1.1 $CARD '' 0.06
$b += Tx 1.25 5.68 5.2 0.35 14 'FFFFFF' 'Trouble replying?' 'l' $true
$b += Tx 1.25 6.02 5.3 0.5 12 $LTG 'Message the host directly and they can add your reply for you.'
$b += Tx 0.9 6.75 8 0.35 12 $MDG 'Event Launch RSVP  |  User Manual  |  v1.0'
$slidesOut += Wrap $b $true

# ---- 3. inject into shell ----
$srcZip = [System.IO.Compression.ZipFile]::OpenRead($shell)
# map slideN -> relId by parsing presentation.xml.rels order
$presEntry = $srcZip.GetEntry('ppt/_rels/presentation.xml.rels')
$sr = New-Object System.IO.StreamReader($presEntry.Open()); $relsXml = $sr.ReadToEnd(); $sr.Close()
$slideRels = [regex]::Matches($relsXml, '<Relationship Id="(rId\d+)"[^>]*slide" Target="slides/slide(\d+)\.xml"')
$relBySlide = @{}
foreach($m in $slideRels){ $relBySlide[$m.Groups[2].Value] = $m.Groups[1].Value }

$ctEntry = $srcZip.GetEntry('[Content_Types].xml')
$sr = New-Object System.IO.StreamReader($ctEntry.Open()); $ctXml = $sr.ReadToEnd(); $sr.Close()
# extend content types for slides 2..8 if the shell only declares some
$extraCT = ''
for($n=1;$n -le 8;$n++){
  if($ctXml -notmatch '/ppt/slides/slide'+$n+'\.xml'){
    $extraCT += '<Override PartName="/ppt/slides/slide'+$n+'.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
  }
}
if($extraCT){ $ctXml = $ctXml.Replace('</Types>', $extraCT+'</Types>') }

# add chart rel + part? dropped chart - using drawn bars instead. skip.

if (Test-Path $outFile) { Remove-Item $outFile -Force }
$fs = [System.IO.File]::Open($outFile, [System.IO.FileMode]::CreateNew)
$out = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create)
foreach($e in $srcZip.Entries){
  if($e.FullName -like 'ppt/slides/slide*.xml' -and $e.FullName -notlike '*.rels'){
    continue   # replaced below
  }
  $ne = $out.CreateEntry($e.FullName)
  $os = $ne.Open(); $is = $e.Open(); $is.CopyTo($os); $os.Close(); $is.Close()
}
# write slide XMLs
$enc = New-Object System.Text.UTF8Encoding $false
for($n=1;$n -le 8;$n++){
  $entry = $out.CreateEntry('ppt/slides/slide'+$n+'.xml')
  $bytes = $enc.GetBytes($slidesOut[$n-1])
  $os = $entry.Open(); $os.Write($bytes,0,$bytes.Length); $os.Close()
}
$out.Dispose(); $fs.Close(); $srcZip.Dispose()
Write-Output ('injected 8 slides into shell -> ' + $outFile)

# ---- 4. verify with real PowerPoint ----
try {
  $pp2 = New-Object -ComObject PowerPoint.Application
  $pres2 = $pp2.Presentations.Open($outFile, $true, $false, $false)
  Write-Output ('FINAL OPENED OK - slides: ' + $pres2.Slides.Count)
  $pres2.Close(); $pp2.Quit()
} catch { Write-Output ('FINAL FAILED: ' + $_.Exception.Message) }
