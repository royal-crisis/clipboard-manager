#Requires AutoHotkey v2.0
#SingleInstance Force

; Initialize GDI+
DllCall("LoadLibrary", "Str", "gdiplus")
si := Buffer(24)
NumPut("UInt", 1, si, 0)
DllCall("gdiplus\GdiplusStartup", "Ptr*", &pToken:=0, "Ptr", si, "Ptr", 0)

; Get Pictures folder path using shell API
GetPicturesPath() {
    ; FOLDERID_Pictures = "{33E28130-4E1E-4676-835A-98395C3BC3BB}"
    FOLDERID_Pictures := Buffer(16)
    DllCall("ole32\CLSIDFromString", "Str", "{33E28130-4E1E-4676-835A-98395C3BC3BB}", "Ptr", FOLDERID_Pictures)
    
    ; SHGetKnownFolderPath
    path := Buffer(A_PtrSize)
    if DllCall("shell32\SHGetKnownFolderPath", "Ptr", FOLDERID_Pictures, "UInt", 0, "Ptr", 0, "Ptr*", &pPath:=0) = 0 {
        path := StrGet(pPath, "UTF-16")
        DllCall("ole32\CoTaskMemFree", "Ptr", pPath)
        return path
    }
    return A_Desktop  ; Fallback to Desktop if Pictures folder not found
}

; Initialize stored clips array
global ClipManager := {
    clips: [],
    maxClips: 25,
    tempDir: A_Temp "\ClipboardManager",
    lastClipboard: "",
    logFile: A_MyDocuments "\ClipboardLog.txt",
    imageDir: GetPicturesPath() "\ClipboardImages"
}

; Create necessary directories
if !DirExist(ClipManager.tempDir)
    DirCreate(ClipManager.tempDir)
if !DirExist(ClipManager.imageDir)
    DirCreate(ClipManager.imageDir)

; Create the main GUI
myGui := Gui("+AlwaysOnTop -Resize")
myGui.Title := "Clipboard History"
myGui.MarginX := 10
myGui.MarginY := 10

; Create ListView with enhanced visual feedback
global lv := myGui.Add("ListView", "x10 y10 w272 h330 -Multi +Grid +LV0x10000", ["Content"])  ; LV0x10000 for hover selection
lv.OnEvent("ItemSelect", ListViewClick)

; Set custom colors for ListView
DllCall("uxtheme\SetWindowTheme", "Ptr", lv.hwnd, "Str", "Explorer", "Ptr", 0)
SendMessage(0x1036, 0, 0xF0F0F0, lv.hwnd)  ; Set hover color (light gray)
SendMessage(0x1026, 0, 0xE0E0E0, lv.hwnd)  ; Set grid color (slightly darker gray)

; Add status text at bottom
global statusText := myGui.Add("Text", "x10 y350 w272 c0x008000", "Ready") ; Green status text

; Add help text at bottom
helpText := myGui.Add("Text", "x10 y380 w272", "Click an item to paste it. Press Win+V to show/hide.")

; Function to store clipboard content
StoreClip(Type) {
    try {
        if Type = 1 { ; Text
            clipText := A_Clipboard
            
            ; Skip if it's empty or the same as last clipboard content
            if (clipText = "" || clipText = ClipManager.lastClipboard)
                return
                
            ClipManager.lastClipboard := clipText
            
            ; Check if this text is already in our list
            for clip in ClipManager.clips {
                if clip.type = "text" && clip.content = clipText
                    return
            }
            
            ; Add new text clip
            ClipManager.clips.InsertAt(1, { type: "text", content: clipText })
            
            ; Log text to file with timestamp
            try {
                FileAppend(FormatTime(A_Now, "yyyy-MM-dd HH:mm:ss") " - " clipText "`n", ClipManager.logFile)
                statusText.Value := "Text copied and logged"
            } catch as e {
                statusText.Value := "Text copied but logging failed"
            }
            statusText.Opt("c0x008000")
        }
        else if Type = 2 { ; Image
            ; Check if clipboard contains a bitmap
            if !DllCall("IsClipboardFormatAvailable", "UInt", 2) { ; CF_BITMAP = 2
                statusText.Value := "No image in clipboard"
                statusText.Opt("c0x800000")
                return
            }
            
            ; Open clipboard
            if !DllCall("OpenClipboard", "Ptr", 0) {
                statusText.Value := "Failed to open clipboard"
                statusText.Opt("c0x800000")
                return
            }
            
            ; Get clipboard bitmap handle
            if !(hBitmap := DllCall("GetClipboardData", "UInt", 2, "Ptr")) {
                DllCall("CloseClipboard")
                statusText.Value := "Failed to get clipboard data"
                statusText.Opt("c0x800000")
                return
            }
            
            ; Get bitmap info
            bi := Buffer(40, 0)
            NumPut("UInt", 40, bi, 0)                    ; Size
            if !DllCall("GetObject", "Ptr", hBitmap, "Int", 40, "Ptr", bi) {
                DllCall("CloseClipboard")
                statusText.Value := "Failed to get bitmap info"
                statusText.Opt("c0x800000")
                return
            }
            
            ; Extract bitmap dimensions
            width := NumGet(bi, 4, "Int")
            height := NumGet(bi, 8, "Int")
            bpp := NumGet(bi, 14, "UShort")
            
            ; Create compatible DC
            if !(hdcScreen := DllCall("GetDC", "Ptr", 0)) {
                DllCall("CloseClipboard")
                statusText.Value := "Failed to get screen DC"
                statusText.Opt("c0x800000")
                return
            }
            
            ; Create memory DC
            if !(hdcMemory := DllCall("CreateCompatibleDC", "Ptr", hdcScreen)) {
                DllCall("ReleaseDC", "Ptr", 0, "Ptr", hdcScreen)
                DllCall("CloseClipboard")
                statusText.Value := "Failed to create memory DC"
                statusText.Opt("c0x800000")
                return
            }
            
            ; Select bitmap into memory DC
            hOldBitmap := DllCall("SelectObject", "Ptr", hdcMemory, "Ptr", hBitmap)
            
            ; Create GDI+ bitmap from DC
            pBitmap := 0
            DllCall("gdiplus\GdipCreateBitmapFromHBITMAP", "Ptr", hBitmap, "Ptr", 0, "Ptr*", &pBitmap)
            
            ; Generate unique filename with timestamp
            timestamp := A_TickCount
            tempFile := ClipManager.tempDir "\" timestamp ".png"  ; Temporary file
            permFile := ClipManager.imageDir "\" timestamp ".png" ; Permanent file in Pictures
            
            ; Save as PNG using GDI+
            ; PNG encoder CLSID: {557CF406-1A04-11D3-9A73-0000F81EF32E}
            CLSID := Buffer(16)
            DllCall("ole32\CLSIDFromString", "Str", "{557CF406-1A04-11D3-9A73-0000F81EF32E}", "Ptr", CLSID)
            
            ; Save bitmap to both locations
            if DllCall("gdiplus\GdipSaveImageToFile", "Ptr", pBitmap, "WStr", tempFile, "Ptr", CLSID, "Ptr", 0) = 0 {
                ; Also save to Pictures folder
                DllCall("gdiplus\GdipSaveImageToFile", "Ptr", pBitmap, "WStr", permFile, "Ptr", CLSID, "Ptr", 0)
                
                ; Check for duplicates
                isDuplicate := false
                for clip in ClipManager.clips {
                    if clip.type = "image" && FileExist(clip.file) {
                        if FileGetSize(clip.file) = FileGetSize(tempFile) {
                            isDuplicate := true
                            break
                        }
                    }
                }
                
                if !isDuplicate {
                    ; Add to clips array
                    ClipManager.clips.InsertAt(1, { type: "image", file: tempFile })
                    statusText.Value := "Image saved to Pictures folder"
                    statusText.Opt("c0x008000")
                } else {
                    FileDelete(tempFile)
                    FileDelete(permFile)
                }
            } else {
                statusText.Value := "Failed to save image"
                statusText.Opt("c0x800000")
            }
            
            ; Cleanup
            DllCall("gdiplus\GdipDisposeImage", "Ptr", pBitmap)
            DllCall("SelectObject", "Ptr", hdcMemory, "Ptr", hOldBitmap)
            DllCall("DeleteDC", "Ptr", hdcMemory)
            DllCall("ReleaseDC", "Ptr", 0, "Ptr", hdcScreen)
            DllCall("CloseClipboard")
        }
        
        ; Remove oldest clip if we exceed maxClips
        while ClipManager.clips.Length > ClipManager.maxClips {
            oldClip := ClipManager.clips.Pop()
            if oldClip.type = "image" && FileExist(oldClip.file)
                FileDelete(oldClip.file)
        }
        
        ; Update ListView
        UpdateListView()
    }
    catch as e {
        OutputDebug("Error in StoreClip: " e.Message "`n" e.Stack)
        statusText.Value := "Error: " e.Message
        statusText.Opt("c0x800000")
    }
}

; Function to update the ListView with clipboard contents
UpdateListView() {
    lv.Delete()
    lv.ModifyCol(1, 260)
    
    for clip in ClipManager.clips {
        if clip.type = "text" {
            displayText := StrLen(clip.content) > 50 ? SubStr(clip.content, 1, 50) "..." : clip.content
            lv.Add(, displayText)
        }
        else if clip.type = "image" {
            ; Extract timestamp from filename (remove path and extension)
            SplitPath(clip.file, &filename, , &ext)
            timestamp := RegExReplace(filename, "\D")  ; Remove non-digits
            
            ; Format the display text
            displayText := timestamp "." ext
            lv.Add(, displayText)
        }
    }
}

; Function to handle ListView clicks
ListViewClick(ctrl, *) {
    row := ctrl.GetNext(0)  ; Get the first selected row
    if row = 0
        return
    
    clip := ClipManager.clips[row]
    
    if clip.type = "text" {
        try {
            A_Clipboard := clip.content
            if ClipWait(0.5) {
                statusText.Value := "Text copied to clipboard"
                statusText.Opt("c0x008000")
                
                ; Hide window before pasting
                WinHide("Clipboard History")
                
                ; Small delay to ensure window is hidden
                Sleep(50)
                
                ; Send paste command
                Send("^v")
            } else {
                statusText.Value := "Failed to copy text"
                statusText.Opt("c0x800000")
            }
        } catch as e {
            statusText.Value := "Error: " e.Message
            statusText.Opt("c0x800000")
        }
    }
    else if clip.type = "image" {
        if FileExist(clip.file) {
            try {
                ; Get screen DC for compatibility
                if !(hdcScreen := DllCall("GetDC", "Ptr", 0))
                    throw Error("Failed to get screen DC")
                
                ; Create a memory DC
                if !(hdcMemory := DllCall("CreateCompatibleDC", "Ptr", hdcScreen))
                    throw Error("Failed to create memory DC")
                
                ; Load the image using GDI+
                pBitmap := 0
                if DllCall("gdiplus\GdipCreateBitmapFromFile", "WStr", clip.file, "Ptr*", &pBitmap) != 0
                    throw Error("Failed to load image")
                
                ; Get the image dimensions
                width := 0, height := 0
                DllCall("gdiplus\GdipGetImageWidth", "Ptr", pBitmap, "UInt*", &width)
                DllCall("gdiplus\GdipGetImageHeight", "Ptr", pBitmap, "UInt*", &height)
                
                ; Create a compatible bitmap
                if !(hBitmap := DllCall("CreateCompatibleBitmap", "Ptr", hdcScreen, "Int", width, "Int", height))
                    throw Error("Failed to create compatible bitmap")
                
                ; Select bitmap into memory DC
                hOldBitmap := DllCall("SelectObject", "Ptr", hdcMemory, "Ptr", hBitmap)
                
                ; Create GDI+ graphics from memory DC
                pGraphics := 0
                if DllCall("gdiplus\GdipCreateFromHDC", "Ptr", hdcMemory, "Ptr*", &pGraphics) != 0
                    throw Error("Failed to create graphics")
                
                ; Draw the image onto the bitmap
                if DllCall("gdiplus\GdipDrawImageI", "Ptr", pGraphics, "Ptr", pBitmap, "Int", 0, "Int", 0) != 0
                    throw Error("Failed to draw image")
                
                ; Open clipboard
                if !DllCall("OpenClipboard", "Ptr", 0)
                    throw Error("Failed to open clipboard")
                
                ; Empty clipboard and set new content
                DllCall("EmptyClipboard")
                
                ; Get the bitmap back from memory DC
                hResultBitmap := DllCall("SelectObject", "Ptr", hdcMemory, "Ptr", hOldBitmap)
                
                ; Copy bitmap to clipboard
                if !DllCall("SetClipboardData", "UInt", 2, "Ptr", hResultBitmap)  ; CF_BITMAP = 2
                    throw Error("Failed to set clipboard data")
                
                statusText.Value := "Image copied to clipboard"
                statusText.Opt("c0x008000")
                
                ; Hide window before pasting
                WinHide("Clipboard History")
                
                ; Longer delay for images
                Sleep(200)
                
                ; Send paste command
                Send("^v")
                
                ; Cleanup GDI+ resources
                DllCall("gdiplus\GdipDeleteGraphics", "Ptr", pGraphics)
                DllCall("gdiplus\GdipDisposeImage", "Ptr", pBitmap)
                
                ; Cleanup GDI resources
                DllCall("DeleteDC", "Ptr", hdcMemory)
                DllCall("ReleaseDC", "Ptr", 0, "Ptr", hdcScreen)
                
                ; Close clipboard
                DllCall("CloseClipboard")
            }
            catch as e {
                ; Cleanup on error
                if IsSet(pGraphics)
                    DllCall("gdiplus\GdipDeleteGraphics", "Ptr", pGraphics)
                if IsSet(pBitmap)
                    DllCall("gdiplus\GdipDisposeImage", "Ptr", pBitmap)
                if IsSet(hdcMemory)
                    DllCall("DeleteDC", "Ptr", hdcMemory)
                if IsSet(hdcScreen)
                    DllCall("ReleaseDC", "Ptr", 0, "Ptr", hdcScreen)
                try DllCall("CloseClipboard")
                
                statusText.Value := "Error: " e.Message
                statusText.Opt("c0x800000")
            }
        } else {
            statusText.Value := "Error: Image file not found"
            statusText.Opt("c0x800000")
        }
    }
}

; Clean up on exit
OnExit(CleanUp)

CleanUp(*) {
    ; Delete all temporary image files
    for clip in ClipManager.clips {
        if clip.type = "image" && FileExist(clip.file)
            FileDelete(clip.file)
    }
    
    ; Shutdown GDI+
    DllCall("gdiplus\GdiplusShutdown", "Ptr", pToken)
}

; Show/hide clipboard manager on Win+V
#v:: {
    if WinExist("Clipboard History") {
        if WinActive() {
            WinHide()
        }
        else {
            WinShow()
            WinActivate()
        }
    }
    else {
        ; Get cursor position
        CoordMode("Mouse", "Screen")
        MouseGetPos(&mouseX, &mouseY)
        
        ; Calculate window position (show near cursor)
        windowWidth := 292
        windowHeight := 410
        
        ; Adjust position to ensure window stays within screen bounds
        pos := AdjustWindowPosition(mouseX, mouseY, windowWidth, windowHeight)
        
        ; Show the GUI at adjusted position
        myGui.Show("x" pos.x " y" pos.y)
    }
}

; Monitor clipboard changes
OnClipboardChange(MonitorClipboard)

MonitorClipboard(Type) {
    if Type = 0  ; Empty clipboard
        return
    
    SetTimer(StoreClip.Bind(Type), -100)  ; Small delay to ensure clipboard content is ready
}

; Function to ensure window stays within monitor boundaries
AdjustWindowPosition(x, y, w, h) {
    ; Get monitor where the cursor is
    monitorIndex := 1
    mouseX := 0
    mouseY := 0
    
    ; Get cursor position
    try {
        mouseX := CoordMode("Mouse", "Screen")
        mouseY := CoordMode("Mouse", "Screen")
    }
    
    ; Get monitor work area
    MonitorGetWorkArea(monitorIndex, &left, &top, &right, &bottom)
    
    ; Adjust position to stay within screen bounds
    if (x + w > right)
        x := right - w
    if (x < left)
        x := left
    if (y + h > bottom)
        y := bottom - h
    if (y < top)
        y := top
    
    return { x: x, y: y }
}
