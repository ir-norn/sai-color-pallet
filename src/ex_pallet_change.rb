#coding:utf-8
#
# ------------------------------------
# 2015 10 17
#
# sai user pallet , save & replace
#
#
#
# ------------------------------------
#
# install lib
# -- gem install inifile
# -- dxruby
#
require 'Win32API'
require 'Win32OLE'
require 'yaml'
require 'dxruby'
require 'inifile'

if ARGV.size != 1
  puts "-- command list --"
  puts "#{File.basename __FILE__} save    # sai pallet img save"
  puts "#{File.basename __FILE__} replace # sai pallet replace"
  exit
end
arg = ARGV[0]

SetCursorPos     = Win32API.new('user32', 'SetCursorPos', ["ii"], "i")
SendInput        = Win32API.new("user32", "SendInput", ["ipi"], "i")
GetSystemMetrics = Win32API.new("user32", "GetSystemMetrics", ["i"], "i")
Screen_w         = GetSystemMetrics.call 0 # SM_CXSCREEN
Screen_h         = GetSystemMetrics.call 1 # SM_CYSCREEN
GetDC            = Win32API.new("user32", "GetDC", ["i"], "i")
TmpDC            = GetDC.call 0
GetPixel         = Win32API.new("gdi32", "GetPixel", ["iii"], "i")
def click key
  _INPUT_MOUSE            = 0
  _MOUSEEVENTF_MOVE       = 0x0001
  _MOUSEEVENTF_LEFTDOWN   = 0x0002
  _MOUSEEVENTF_LEFTUP     = 0x0004
  _MOUSEEVENTF_RIGTHDOWN  = 0x0008
  _MOUSEEVENTF_RIGTHUP    = 0x0010
  #
  _MOUSEEVENTF_MIDDLEDOWN = 0x0020
  _MOUSEEVENTF_MIDDLEUP   = 0x0040
  _MOUSEEVENTF_WHEEL      = 0x0800
  _MOUSEEVENTF_ABSOLUTE   = 0x8000
  in1 = 0
  in2 = 0
  case key
  when :left
    in1 = [_INPUT_MOUSE, 0, 0, 0, _MOUSEEVENTF_LEFTDOWN,   0, 0].pack('LllLLLL')
    in2 = [_INPUT_MOUSE, 0, 0, 0, _MOUSEEVENTF_LEFTUP,     0, 0].pack('LllLLLL')
  when :right
    in1 = [_INPUT_MOUSE, 0, 0, 0, _MOUSEEVENTF_RIGTHDOWN,  0, 0].pack('LllLLLL')
    in2 = [_INPUT_MOUSE, 0, 0, 0, _MOUSEEVENTF_RIGTHUP,    0, 0].pack('LllLLLL')
  when :middle
    in1 = [_INPUT_MOUSE, 0, 0, 0, _MOUSEEVENTF_MIDDLEDOWN, 0, 0].pack('LllLLLL')
    in2 = [_INPUT_MOUSE, 0, 0, 0, _MOUSEEVENTF_MIDDLEUP,   0, 0].pack('LllLLLL')
  end
  input = in1 + in2
  SendInput.call 2, input , in1.length
end
#
# todo this GetPixel_rb function Hex 6 moji return
# now 6 moji ika return
# but "000044" is "44" 
#
def GetPixel_rb x , y
  GetPixel.call( TmpDC , x , y ).to_s(16).split(/(..)/).reverse.join
end
# ---- misic.ini 読み込み
#
# ------- 設定例 -------
#
# [Swatch]
#
# ;========================================================================================
# ; カラーパレットのカスタマイズ
# ;========================================================================================
# ;
# ; 下記の設定を有効にするとα版までの 16×8 マスの表示になります。
# ; (各行頭の";"を削除することで有効になります)
# ;
# Size      = 13     ; パレットのマスの大きさ
# Cols      = 12     ; 横方向のマスの数
# Rows      = 24      ; 縦方向のマスの数
# Width     = 300    ; パレットウィンドウの幅
# Height    = 600     ; パレットウィンドウの高さ
# ShowHSB   = 1      ; 水平スクロールバーを表示するかどうか (0=表示しない 1=表示する)
# ShowVSB   = 1      ; 垂直スクロールバーを表示するかどうか (0=表示しない 1=表示する)
#
#
# ini = IniFile.load('C:/soft/SAI_c/misc.ini')
# Size = ini["Swatch"]["Size"]
# Rows = ini["Swatch"]["Rows"]
# Cols = ini["Swatch"]["Cols"]
# puts "Size #{Size}"
# puts "Rows #{Rows}"
# puts "Cols #{Cols}"
# (puts "ini file load error" ; exit) if not ini
#
# photo shop x , y ## memo
# 1 33
#  19 51
# ------------------------------
#
#  手順 save
#
#  1. saiのウィンドウを最大化 （ F11ではないほう
#  2. 他のウィンドウを全て閉じてユーザーパレットのみを表示する
#  3.　↑左側へ表示
#  4. sai画面手前にした状態でこのスクリプトを実行
#  5. ruby sai_pallet_change.rb dump
#
#  パレットコピーの sai_color.bmp が生成されれば成功
#
#
#
#
# ----- sai pallet img save -----------------------
$debug_mode = false   ;;puts $debug_mode && "debug_mode"
#
save = -> do  
  ini = IniFile.load('C:/soft/SAI_c/misc.ini')
  size = ini["Swatch"]["Size"]
  rows = ini["Swatch"]["Rows"]
  cols = ini["Swatch"]["Cols"]
  cols = rows = 5 if $debug_mode
  # sai absolute
  init_x = 6
  init_y = 75
  v =
  cols.times.map do | a |
    rows.times.map do | b |
      x = init_x  + (size / 2) + (size * a) # absolute position left 6 top 75   ...not padding 1?
      y = init_y  + (size / 2) + (size * b)
      [ x , y ]
    end end.flatten.each_slice(2).to_a
  
  colors =
  v.map do | x , y |
    SetCursorPos.call x , y
    sleep 0.01    
    if (c = GetPixel_rb( x , y )) == "f0f0f0"
       if GetPixel_rb( x-2, y+2 ) == "e0e0e0"
        next
      end
    end
#    c = (c+"000000")[0..5]  ## min patch
    print "#" , c , " "
    [ x , y , c  ]
  end.compact

  d =
  colors.map do | _x , _y , c |
    ## min patch 2015 10 22
    r = (c[0..1]||0).to_i(16)
    g = (c[2..3]||0).to_i(16)
    b = (c[4..5]||0).to_i(16)
    Image.new( size-1 , size-1 , [255, r , g , b ] )
  end

  count = 0
  Window.height = 1000
  Window.loop do
    d.each_with_index do |a,i|
      Window.draw( colors[i][0] - size / 2 , colors[i][1] - size / 2 , a )
    end
    if (count+=1) > 120 #wait
      Window.getScreenShot( "sai_color.bmp" )
      break
    end
  end
end

#
#  手順 replace
#
#  1. ユーザーパレットを上書きしたいsaiを起動する
#  2. saveの手順と同じ
#  3. saveの手順と同じ
#  4. sai_color.bmpを開く 　※　開いた後に縮小・拡大・移動等の一切の操作をしない事
#  5. ruby sai_pallet_change.rb replace
#  6. マウスとキーボードによるマクロ処理が行われるので処理終わるまでpc放置
#
#
#
# ----- sai pallet replace  -----------------------


rep = -> do
# sai sbsolute
  ini = IniFile.load('C:/soft/SAI_c/misc.ini')
  size = ini["Swatch"]["Size"]
  rows = ini["Swatch"]["Rows"]
  cols = ini["Swatch"]["Cols"]
  cols = rows = 5 if $debug_mode
  img_cell_size = 19 - 1 + ( pad = 2 ) # .img colors absolute & size
#  rows = 25
#  cols = 10
#  puts ".img Size #{size} _ Rows #{rows} _ Cols #{cols}"
  xy =
  Screen_h.times.each do | i |
    if GetPixel_rb( Screen_w / 2 , 0 + i )  == "0"
      break p [ Screen_w / 2 , 0 + i ]
    end end
  xy =
  Screen_w.times do | i |
    if GetPixel_rb( xy[0] - i , xy[1] ) != "0"
      break p [ xy[0] - i  +  1 , xy[1]]
    end end
  xy =
  300.times do | i |
    if GetPixel_rb( xy[0] + img_cell_size/2 , xy[1] + i ) != "0"
      break p [ xy[0]  , xy[1] + i ] 
    end end
  init_x = xy[0]
  init_y = xy[1]

  vv = 
  cols.times.map do | a |
    rows.times.map do | b |
      x = init_x + (size / 2) + (size * a)
      y = init_y + (size / 2) + (size * b)
      [ x , y ]
    end end.flatten.each_slice(2).to_a
 # p vv
  puts "Size #{size} _ Rows #{rows} _ Cols #{cols} _ img_cell_size #{img_cell_size}"
  # sai absolute position
  init_x = 6
  init_y = 75
  v =
  cols.times.map do | a |
    rows.times.map do | b |
      x = init_x  + (size / 2) + (size * a) # absolute position left 6 top 75   ...not padding 1?
      y = init_y  + (size / 2) + (size * b)
      [ x , y ]
    end end.flatten.each_slice(2).to_a

  wsh = WIN32OLE.new('WScript.Shell')
  vv_tmp = vv.clone
  v.map do | x , y |
    img_x , img_y = vv_tmp.shift
    SetCursorPos.call img_x , img_y 
    sleep 0.1 ; click :right # not active window __ todo
    sleep 0.1 ; click :right

    # canvas color get
    case GetPixel_rb( img_x , img_y )
      when "0"      then next
      when "ffffff" then next
      when "e0e0e0" then next
      when "f0f0f0" then next
    end

    # replace
    SetCursorPos.call x , y
    sleep 0.1 ; click :right
    sleep 0.2 ; wsh.Sendkeys("{DOWN}")
    sleep 0.2 ; wsh.Sendkeys("{DOWN}")
    sleep 0.2 ; wsh.Sendkeys("{ENTER}")
  end
end

#--------------------
#
#
#
#
#
#
#
#
# command line
#
#
#
# --------------------------------------
case arg
  when "save"
    save.call
  when "replace"
    rep.call
  else
    puts "command err"
end
