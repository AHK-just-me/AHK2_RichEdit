#DllLoad "Comdlg32.dll"
Class RichEditDlgs {
   Static Call(*) => False
   ; ===================================================================================================================
   ; ===================================================================================================================
   ; RICHEDIT COMMON DIALOGS ===========================================================================================
   ; ===================================================================================================================
   Static FindReplMsg := DllCall("RegisterWindowMessage", "Str", "commdlg_FindReplace", "UInt") ; FINDMSGSTRING
   ; ===================================================================================================================
   ; Most of the following methods are based on DLG 5.01 by majkinetor
   ; http://www.autohotkey.com/board/topic/15836-module-dlg-501/
   ; ===================================================================================================================
   Static ChooseColor(RE, Color := "") { ; Choose color dialog box
   ; ===================================================================================================================
      ; RE : RichEdit object
      Static CC_Size := A_PtrSize * 9, CCU := Buffer(64, 0)
      GuiHwnd := RE.Gui.Hwnd
      If (Color != "")
         Color := RE.GetBGR(Color)
      Else
         Color := 0x000000
      CC :=  Buffer(CC_Size, 0)                    ; CHOOSECOLOR structure
      NumPut("UInt", CC_Size, CC, 0)               ; lStructSize
      NumPut("UPtr", GuiHwnd, CC, A_PtrSize)       ; hwndOwner makes dialog modal
      NumPut("UInt", Color, CC, A_PtrSize * 3)     ; rgbResult
      NumPut("UPtr", CCU.Ptr, CC, A_PtrSize * 4)   ; COLORREF *lpCustColors (16)
      NumPut("UInt", 0x0101, CC, A_PtrSize * 5)    ; Flags: CC_ANYCOLOR | CC_RGBINIT | ; CC_FULLOPEN
      R := DllCall("Comdlg32.dll\ChooseColor", "Ptr", CC.Ptr, "UInt")
      Return (R = 0) ? "" : RE.GetRGB(NumGet(CC, A_PtrSize * 3, "UInt"))
   }
   ; ===================================================================================================================
   Static ChooseFont(RE) { ; Choose font dialog box
   ; ===================================================================================================================
      ; RE : RichEdit object
      DC := DllCall("GetDC", "Ptr", RE.Gui.Hwnd, "Ptr")
      LP := DllCall("GetDeviceCaps", "Ptr", DC, "UInt", 90, "Int")   ; LOGPIXELSY
      DllCall("ReleaseDC", "Ptr", RE.Gui.Hwnd, "Ptr", DC)
      ; Get current font
      Font := RE.GetFont()
      ; LF_FACENAME = 32
      LF := Buffer(92, 0)                   ; LOGFONT structure
      Size := -(Font.Size * LP / 72)
      NumPut("Int", Size, LF, 0)            ; lfHeight
      If InStr(Font.Style, "B")
         NumPut("Int", 700, LF, 16)         ; lfWeight
      If InStr(Font.Style, "I")
         NumPut("UChar", 1, LF, 20)         ; lfItalic
      If InStr(Font.Style, "U")
         NumPut("UChar", 1, LF, 21)         ; lfUnderline
      If InStr(Font.Style, "S")
         NumPut("UChar", 1, LF, 22)         ; lfStrikeOut
      NumPut("UChar", Font.CharSet, LF, 23) ; lfCharSet
      StrPut(Font.Name, LF.Ptr + 28, 32)
      ; CF_BOTH = 3, CF_INITTOLOGFONTSTRUCT = 0x40, CF_EFFECTS = 0x100, CF_SCRIPTSONLY = 0x400
      ; CF_NOVECTORFONTS = 0x800, CF_NOSIMULATIONS = 0x1000, CF_LIMITSIZE = 0x2000, CF_WYSIWYG = 0x8000
      ; CF_TTONLY = 0x40000, CF_FORCEFONTEXIST =0x10000, CF_SELECTSCRIPT = 0x400000
      ; CF_NOVERTFONTS =0x01000000
      Flags := 0x00002141 ; 0x01013940
      If (Font.Color = "Auto")
         Color := DllCall("GetSysColor", "Int", 8, "UInt") ; COLOR_WINDOWTEXT = 8
      Else
         Color := RE.GetBGR(Font.Color)
      CF_Size := (A_PtrSize = 8 ? (A_PtrSize * 10) + (4 * 4) + A_PtrSize : (A_PtrSize * 14) + 4)
      CF := Buffer(CF_Size, 0)                           ; CHOOSEFONT structure
      NumPut("UInt", CF_Size, CF)                        ; lStructSize
      NumPut("UPtr", RE.Gui.Hwnd, CF, A_PtrSize)	      ; hwndOwner (makes dialog modal)
      NumPut("UPtr", LF.Ptr, CF, A_PtrSize * 3)	         ; lpLogFont
      NumPut("UInt", Flags, CF, (A_PtrSize * 4) + 4)     ; Flags
      NumPut("UInt", Color, CF, (A_PtrSize * 4) + 8)     ; rgbColors
      OffSet := (A_PtrSize = 8 ? (A_PtrSize * 11) + 4 : (A_PtrSize * 12) + 4)
      NumPut("Int", 4, CF, Offset)                       ; nSizeMin
      NumPut("Int", 160, CF, OffSet + 4)                 ; nSizeMax
      ; Call ChooseFont Dialog
      If !DllCall("Comdlg32.dll\ChooseFont", "Ptr", CF.Ptr, "UInt")
         Return false
      ; Get name
      Font.Name := StrGet(LF.Ptr + 28, 32)
   	; Get size
   	Font.Size := NumGet(CF, A_PtrSize * 4, "Int") / 10
      ; Get styles
   	Font.Style := ""
   	If NumGet(LF, 16, "Int") >= 700
   	   Font.Style .= "B"
   	If NumGet(LF, 20, "UChar")
         Font.Style .= "I"
   	If NumGet(LF, 21, "UChar")
         Font.Style .= "U"
   	If NumGet(LF, 22, "UChar")
         Font.Style .= "S"
      OffSet := A_PtrSize * (A_PtrSize = 8 ? 11 : 12)
      FontType := NumGet(CF, Offset, "UShort")
      If (FontType & 0x0100) && !InStr(Font.Style, "B") ; BOLD_FONTTYPE
         Font.Style .= "B"
      If (FontType & 0x0200) && !InStr(Font.Style, "I") ; ITALIC_FONTTYPE
         Font.Style .= "I"
      If (Font.Style = "")
         Font.Style := "N"
      ; Get character set
      Font.CharSet := NumGet(LF, 23, "UChar")
      ; We don't use the limited colors of the font dialog
      ; Return selected values
      Return RE.SetFont(Font)
   }
   ; ===================================================================================================================
   Static FileDlg(RE, Mode, File := "") { ; Open and save as dialog box
   ; ===================================================================================================================
      ; RE   : RichEdit object
      ; Mode : O = Open, S = Save
      ; File : optional file name
   	Static OFN_ALLOWMULTISELECT := 0x200,    OFN_EXTENSIONDIFFERENT := 0x400, OFN_CREATEPROMPT := 0x2000,
             OFN_DONTADDTORECENT := 0x2000000, OFN_FILEMUSTEXIST := 0x1000,     OFN_FORCESHOWHIDDEN := 0x10000000,
             OFN_HIDEREADONLY := 0x4,          OFN_NOCHANGEDIR := 0x8,          OFN_NODEREFERENCELINKS := 0x100000,
             OFN_NOVALIDATE := 0x100,          OFN_OVERWRITEPROMPT := 0x2,      OFN_PATHMUSTEXIST := 0x800,
             OFN_READONLY := 0x1,              OFN_SHOWHELP := 0x10,            OFN_NOREADONLYRETURN := 0x8000,
             OFN_NOTESTFILECREATE := 0x10000,  OFN_ENABLEXPLORER := 0x80000
             OFN_Size := (4 * 5) + (2 * 2) + (A_PtrSize * 16)
      Static FilterN1 := "RichText",   FilterP1 := "*.rtf",
             FilterN2 := "Text",       FilterP2 := "*.txt",
             FilterN3 := "AutoHotkey", FilterP3 := "*.ahk",
             DefExt := "rtf",
             DefFilter := 1
   	SplitPath(File, &Name := "", &Dir := "")
      Flags := OFN_ENABLEXPLORER
      Flags |= Mode = "O" ? OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST | OFN_HIDEREADONLY
                          : OFN_OVERWRITEPROMPT
   	VarSetStrCapacity(&FileName, 512)
      FileName := Name
   	LenN1 := (StrLen(FilterN1) + 1) * 2, LenP1 := (StrLen(FilterP1) + 1) * 2
   	LenN2 := (StrLen(FilterN2) + 1) * 2, LenP2 := (StrLen(FilterP2) + 1) * 2
   	LenN3 := (StrLen(FilterN3) + 1) * 2, LenP3 := (StrLen(FilterP3) + 1) * 2
      Filter := Buffer(LenN1 + LenP1 + LenN2 + LenP2 + LenN3 + LenP3 + 4, 0)
      Adr := Filter.Ptr
      StrPut(FilterN1, Adr)
      StrPut(FilterP1, Adr += LenN1)
      StrPut(FilterN2, Adr += LenP1)
      StrPut(FilterP2, Adr += LenN2)
      StrPut(FilterN3, Adr += LenP2)
      StrPut(FilterP3, Adr += LenN3)
      OFN := Buffer(OFN_Size, 0)                     ; OPENFILENAME Structure
   	NumPut("UInt", OFN_Size, OFN, 0)
      Offset := A_PtrSize
   	NumPut("Ptr", RE.Gui.Hwnd, OFN, Offset)        ; HWND owner
      Offset += A_PtrSize * 2
   	NumPut("Ptr", Filter.Ptr, OFN, OffSet)         ; Pointer to FilterStruc
      OffSet += (A_PtrSize * 2) + 4
      OffFilter := Offset
   	NumPut("UInt", DefFilter, OFN, Offset)         ; DefaultFilter Pair
      OffSet += 4
   	NumPut("Ptr", StrPtr(FileName), OFN, OffSet)   ; lpstrFile / InitialisationFileName
      Offset += A_PtrSize
   	NumPut("UInt", 512, OFN, Offset)               ; MaxFile / lpstrFile length
      OffSet += A_PtrSize * 3
   	NumPut("Ptr", StrPtr(Dir), OFN, Offset)        ; StartDir
      Offset += A_PtrSize * 2
   	NumPut("UInt", Flags, OFN, Offset)             ; Flags
      Offset += 8
   	NumPut("Ptr", StrPtr(DefExt), OFN, Offset)     ; DefaultExt
      R := Mode = "S" ? DllCall("Comdlg32.dll\GetSaveFileNameW", "Ptr", OFN.Ptr, "UInt")
                      : DllCall("Comdlg32.dll\GetOpenFileNameW", "Ptr", OFN.Ptr, "UInt")
   	If !(R)
         Return ""
      DefFilter := NumGet(OFN, OffFilter, "UInt")
   	Return StrGet(StrPtr(FileName))
   }
   ; ===================================================================================================================
   Static FindText(RE) { ; Find dialog box
   ; ===================================================================================================================
      ; RE : RichEdit object
   	Static FR_DOWN := 1, FR_MATCHCASE := 4, FR_WHOLEWORD := 2,
   	       Buf := "", BufLen := 256, FR := "", FR_Size := A_PtrSize * 10
      Text := RE.GetSelText()
      Buf := ""
      VarSetStrCapacity(&Buf, BufLen)
      If (Text != "") && !RegExMatch(Text, "\W")
         Buf := Text
      FR := Buffer(FR_Size, 0)
   	NumPut("UInt", FR_Size, FR)
      Offset := A_PtrSize
   	NumPut("UPtr", RE.Gui.Hwnd, FR, Offset)  ; hwndOwner
      OffSet += A_PtrSize * 2
   	NumPut("UInt", FR_DOWN, FR, Offset)	     ; Flags
      OffSet += A_PtrSize
   	NumPut("UPtr", StrPtr(Buf), FR, Offset)  ; lpstrFindWhat
      OffSet += A_PtrSize * 2
   	NumPut("Short", BufLen,	FR, Offset)      ; wFindWhatLen
      This.FindTextProc("Init", RE.HWND, "")
   	OnMessage(RichEditDlgs.FindReplMsg, RichEditDlgs.FindTextProc)
   	Return DllCall("Comdlg32.dll\FindTextW", "Ptr", FR.Ptr, "UPtr")
   }
   ; -------------------------------------------------------------------------------------------------------------------
   Static FindTextProc(L, M, H) { ; skipped wParam, can be found in "This" when called by system
      ; Find dialog callback procedure
      ; EM_FINDTEXTEXW = 0x047C, EM_EXGETSEL = 0x0434, EM_EXSETSEL = 0x0437, EM_SCROLLCARET = 0x00B7
      ; FR_DOWN = 1, FR_WHOLEWORD = 2, FR_MATCHCASE = 4,
   	Static FR_DOWN := 1, FR_MATCHCASE := 4, FR_WHOLEWORD := 2 , FR_FINDNEXT := 0x8, FR_DIALOGTERM := 0x40,
             HWND := 0
      If (L = "Init") {
         HWND := M
         Return True
      }
      Flags := NumGet(L, A_PtrSize * 3, "UInt")
      If (Flags & FR_DIALOGTERM) {
         OnMessage(RichEditDlgs.FindReplMsg, RichEditDlgs.FindTextProc, 0)
         If (RE := GuiCtrlFromHwnd(HWND))
            RE.Focus()
         HWND := 0
         Return
      }
      CR := Buffer(8, 0)
      SendMessage(0x0434, 0, CR.Ptr, HWND)
      Min := (Flags & FR_DOWN) ? NumGet(CR, 4, "Int") : NumGet(CR, 0, "Int")
      Max := (Flags & FR_DOWN) ? -1 : 0
      OffSet := A_PtrSize * 4
      Find := StrGet(NumGet(L, Offset, "UPtr"))
      FTX := Buffer(16 + A_PtrSize, 0)
      NumPut("Int", Min, "Int", Max, "UPtr", StrPtr(Find), FTX)
      SendMessage(0x047C, Flags, FTX.Ptr, HWND)
      S := NumGet(FTX, 8 + A_PtrSize, "Int"), E := NumGet(FTX, 12 + A_PtrSize, "Int")
      If (S = -1) && (E = -1)
         MsgBox("No (further) occurence found!", "Find", 262208)
      Else {
         SendMessage(0x0437, 0, FTX.Ptr + 8 + A_PtrSize, HWND)
         SendMessage(0x00B7, 0, 0, HWND)
      }
   }
   ; ===================================================================================================================
   Static PageSetup(RE) { ; Page setup dialog box
   ; ===================================================================================================================
      ; RE : RichEdit object
      ; http://msdn.microsoft.com/en-us/library/ms646842(v=vs.85).aspx
      Static PSD_DEFAULTMINMARGINS             := 0x00000000, ; default (printer's)
             PSD_INWININIINTLMEASURE           := 0x00000000, ; 1st of 4 possible
             PSD_MINMARGINS                    := 0x00000001, ; use caller's
             PSD_MARGINS                       := 0x00000002, ; use caller's
             PSD_INTHOUSANDTHSOFINCHES         := 0x00000004, ; 2nd of 4 possible
             PSD_INHUNDREDTHSOFMILLIMETERS     := 0x00000008, ; 3rd of 4 possible
             PSD_DISABLEMARGINS                := 0x00000010,
             PSD_DISABLEPRINTER                := 0x00000020,
             PSD_NOWARNING                     := 0x00000080, ; must be same as PD_*
             PSD_DISABLEORIENTATION            := 0x00000100,
             PSD_RETURNDEFAULT                 := 0x00000400, ; must be same as PD_*
             PSD_DISABLEPAPER                  := 0x00000200,
             PSD_SHOWHELP                      := 0x00000800, ; must be same as PD_*
             PSD_ENABLEPAGESETUPHOOK           := 0x00002000, ; must be same as PD_*
             PSD_ENABLEPAGESETUPTEMPLATE       := 0x00008000, ; must be same as PD_*
             PSD_ENABLEPAGESETUPTEMPLATEHANDLE := 0x00020000, ; must be same as PD_*
             PSD_ENABLEPAGEPAINTHOOK           := 0x00040000,
             PSD_DISABLEPAGEPAINTING           := 0x00080000,
             PSD_NONETWORKBUTTON               := 0x00200000, ; must be same as PD_*
             I := 1000, ; thousandth of inches
             M := 2540, ; hundredth of millimeters
             Margins := {},
             Metrics := "",
             PSD_Size := (4 * 10) + (A_PtrSize * 11),
             PD_Size := (A_PtrSize = 8 ? (13 * A_PtrSize) + 16 : 66),
             OffFlags := 4 * A_PtrSize,
             OffMargins := OffFlags + (4 * 7)
      PSD := Buffer(PSD_Size, 0)                    ; PAGESETUPDLG structure
      NumPut("UInt", PSD_Size, PSD)
      NumPut("UPtr", RE.Gui.Hwnd, PSD, A_PtrSize)   ; hwndOwner
      Flags := PSD_MARGINS | PSD_DISABLEPRINTER | PSD_DISABLEORIENTATION | PSD_DISABLEPAPER
      NumPut("Int", Flags, PSD, OffFlags)           ; Flags
      Offset := OffMargins
      NumPut("Int", RE.Margins.L, PSD, Offset += 0) ; rtMargin left
      NumPut("Int", RE.Margins.T, PSD, Offset += 4) ; rtMargin top
      NumPut("Int", RE.Margins.R, PSD, Offset += 4) ; rtMargin right
      NumPut("Int", RE.Margins.B, PSD, Offset += 4) ; rtMargin bottom
      If !DllCall("Comdlg32.dll\PageSetupDlg", "Ptr", PSD.Ptr, "UInt")
         Return False
      DllCall("Kernel32.dll\GlobalFree", "Ptr", NumGet(PSD, 2 * A_PtrSize, "UPtr"))
      DllCall("Kernel32.dll\GlobalFree", "Ptr", NumGet(PSD, 3 * A_PtrSize, "UPtr"))
      Flags := NumGet(PSD, OffFlags, "UInt")
      Metrics := (Flags & PSD_INTHOUSANDTHSOFINCHES) ? I : M
      Offset := OffMargins
      RE.Margins.L := NumGet(PSD, Offset += 0, "Int")
      RE.Margins.T := NumGet(PSD, Offset += 4, "Int")
      RE.Margins.R := NumGet(PSD, Offset += 4, "Int")
      RE.Margins.B := NumGet(PSD, Offset += 4, "Int")
      RE.Margins.LT := Round((RE.Margins.L / Metrics) * 1440) ; Left as twips
      RE.Margins.TT := Round((RE.Margins.T / Metrics) * 1440) ; Top as twips
      RE.Margins.RT := Round((RE.Margins.R / Metrics) * 1440) ; Right as twips
      RE.Margins.BT := Round((RE.Margins.B / Metrics) * 1440) ; Bottom as twips
      Return True
   }
   ; ===================================================================================================================
   Static ReplaceText(RE) { ; Replace dialog box
   ; ===================================================================================================================
      ; RE : RichEdit object
   	Static FR_DOWN := 1, FR_MATCHCASE := 4, FR_WHOLEWORD := 2,
   	       FBuf := "", RBuf := "", BufLen := 256, FR := "", FR_Size := A_PtrSize * 10
      Text := RE.GetSelText()
      FBuf := RBuf := ""
      VarSetStrCapacity(&FBuf, BufLen)
      If (Text != "") && !RegExMatch(Text, "\W")
         FBuf := Text
      VarSetStrCapacity(&RBuf, BufLen)
      FR := Buffer(FR_Size, 0)
   	NumPut("UInt", FR_Size, FR)
      Offset := A_PtrSize
   	NumPut("UPtr", RE.Gui.Hwnd, FR, Offset)              ; hwndOwner
      OffSet += A_PtrSize * 2
   	NumPut("UInt", FR_DOWN, FR, Offset)	                 ; Flags
      OffSet += A_PtrSize
   	NumPut("UPtr", StrPtr(FBuf), FR, Offset)             ; lpstrFindWhat
      OffSet += A_PtrSize
   	NumPut("UPtr", StrPtr(RBuf), FR, Offset)             ; lpstrReplaceWith
      OffSet += A_PtrSize
   	NumPut("Short", BufLen,	"Short", BufLen, FR, Offset) ; wFindWhatLen, wReplaceWithLen
      This.ReplaceTextProc("Init", RE.HWND, "")
   	OnMessage(RichEditDlgs.FindReplMsg, RichEditDlgs.ReplaceTextProc)
   	Return DllCall("Comdlg32.dll\ReplaceText", "Ptr", FR.Ptr, "UPtr")
   }
   ; -------------------------------------------------------------------------------------------------------------------
   Static ReplaceTextProc(L, M, H) { ; skipped wParam, can be found in "This" when called by system
      ; Replace dialog callback procedure
      ; EM_FINDTEXTEXW = 0x047C, EM_EXGETSEL = 0x0434, EM_EXSETSEL = 0x0437
      ; EM_REPLACESEL = 0xC2, EM_SCROLLCARET = 0x00B7
      ; FR_DOWN = 1, FR_WHOLEWORD = 2, FR_MATCHCASE = 4,
   	Static FR_DOWN := 1, FR_MATCHCASE := 4, FR_WHOLEWORD := 2, FR_FINDNEXT := 0x8,
             FR_REPLACE := 0x10, FR_REPLACEALL := 0x20, FR_DIALOGTERM := 0x40,
             HWND := 0, Min := "", Max := "", FS := "", FE := "",
             OffFind := A_PtrSize * 4, OffRepl := A_PtrSize * 5
      If (L = "Init") {
         HWND := M, FS := "", FE := ""
         Return True
      }
      Flags := NumGet(L, A_PtrSize * 3, "UInt")
      If (Flags & FR_DIALOGTERM) {
         OnMessage(RichEditDlgs.FindReplMsg, RichEditDlgs.ReplaceTextProc, 0)
         If (RE := GuiCtrlFromHwnd(HWND))
            RE.Focus()
         HWND := 0
         Return
      }
      If (Flags & FR_REPLACE) {
         IF (FS >= 0) && (FE >= 0) {
            SendMessage(0xC2, 1, NumGet(L, OffRepl, "UPtr"), HWND)
            Flags |= FR_FINDNEXT
         }
         Else
            Return
      }
      If (Flags & FR_FINDNEXT) {
         CR := Buffer(8, 0)
         SendMessage(0x0434, 0, CR.Ptr, HWND)
         Min := NumGet(CR, 4)
         FS := FE := ""
         Find := NumGet(L, OffFind, "UPtr")
         FTX := Buffer(16 + A_PtrSize, 0)
         NumPut("Int", Min, "Int", -1, "Ptr", Find, FTX)
         SendMessage(0x047C, Flags, FTX.Ptr, HWND)
         S := NumGet(FTX, 8 + A_PtrSize, "Int"), E := NumGet(FTX, 12 + A_PtrSize, "Int")
         If (S = -1) && (E = -1)
            MsgBox("No (further) occurence found!", "Replace", 262208)
         Else {
            SendMessage(0x0437, 0, FTX.Ptr + 8 + A_PtrSize, HWND)
            SendMessage(0x00B7, 0, 0, HWND)
            FS := S, FE := E
         }
         Return
      }
      If (Flags & FR_REPLACEALL) {
         CR := Buffer(8, 0)
         SendMessage(0x0434, 0, CR.Ptr, HWND)
         If (FS = "")
            FS := FE := 0
         DllCall("User32.dll\LockWindowUpdate", "Ptr", HWND)
         Find := NumGet(L, OffFind, "UPtr")
         FTX := Buffer(16 + A_PtrSize, 0)
         NumPut("Int", FS, "Int", -1, "Ptr",, Find, FTX)
         While (FS >= 0) && (FE >= 0) {
            SendMessage(0x044F, Flags, FTX.Ptr, HWND)
            FS := NumGet(FTX, A_PtrSize + 8, "Int"), FE := NumGet(FTX, A_PtrSize + 12, "Int")
            If (FS >= 0) && (FE >= 0) {
               SendMessage(0x0437, 0, FTX.Ptr + 8 + A_PtrSize, HWND)
               SendMessage(0xC2, 1, NumGet(L + 0, OffRepl, "UPtr" ), HWND)
               NumPut("Int", FE, FTX)
            }
         }
         SendMessage(0x0437, 0, CR.Ptr, HWND)
         DllCall("User32.dll\LockWindowUpdate", "Ptr", 0)
         Return
      }
   }
}
