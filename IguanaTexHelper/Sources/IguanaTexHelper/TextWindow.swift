import AppKit

class TextWindow: NSWindow {
    override var canBecomeKey: Bool { true }
    override var parent: NSWindow? {
        willSet {
            guard parent != nil else { return }
            NotificationCenter.default.removeObserver(
                self,
                name: NSWindow.didBecomeMainNotification,
                object: parent
            )
        }
        didSet {
            guard parent != nil else { return }
            NotificationCenter.default.addObserver(
                self,
                selector: #selector(onParentDidBecomeMain(_:)),
                name: NSWindow.didBecomeMainNotification,
                object: parent
            )
        }
    }

    weak var resizeTarget: AnyObject? = nil
    let scrollView: NSScrollView = NSScrollView()
    let textView: NSTextView = NSTextView()
    var wordWrap: Bool = true { didSet { setupWordWrap() } }

    init() {
        super.init(
            contentRect: NSZeroRect,
            styleMask: .borderless,
            backing: .buffered,
            defer: false
        )
        setupViews()
        setupWordWrap()
    }

    func setupViews() {
        textView.allowsUndo = true
        textView.isRichText = false
        textView.font = NSFont.userFixedPitchFont(ofSize: 10)
        textView.isAutomaticDashSubstitutionEnabled = false
        textView.isAutomaticQuoteSubstitutionEnabled = false
        textView.maxSize = NSSize(
            width: CGFloat.greatestFiniteMagnitude,
            height: CGFloat.greatestFiniteMagnitude
        )
        textView.isVerticallyResizable = true
        textView.isHorizontallyResizable = true
        textView.autoresizingMask = [.width, .height]

        scrollView.borderType = .bezelBorder
        scrollView.hasVerticalScroller = true
        scrollView.autohidesScrollers = true
        scrollView.autoresizingMask = [.width, .height]

        scrollView.documentView = textView
        self.contentView = scrollView
        self.initialFirstResponder = textView
    }

    func setupWordWrap() {
        scrollView.hasHorizontalScroller = wordWrap
        textView.textContainer?.containerSize = NSSize(
            width: wordWrap ? scrollView.contentSize.width : CGFloat.greatestFiniteMagnitude,
            height: CGFloat.greatestFiniteMagnitude
        )
        textView.textContainer?.widthTracksTextView = wordWrap
        textView.sizeToFit()
    }

    @objc func onParentDidBecomeMain(_ notificaton: Notification) {
        self.orderFront(nil)
    }
}

var textWindows: [Int64: TextWindow] = [:]
var lastHandle: Int64 = 0

@_cdecl("TWInit")
public func TWInit() -> Int64 {
    let handle = lastHandle + 1
    lastHandle += 1
    let window = TextWindow()
    window.isReleasedWhenClosed = false
    textWindows[handle] = window
    return handle
}

@_cdecl("TWTerm")
public func TWTerm(_ handle: Int64, _: Int64, _: Int64, _: Int64) -> Int64 {
    guard let window = textWindows[handle] else { return 0 }
    window.close()
    textWindows[handle] = nil
    return 0
}

@_cdecl("TWShow")
public func TWShow(_ handle: Int64, _: Int64, _: Int64, _: Int64) -> Int64 {
    guard let window = textWindows[handle] else { return 0 }
    guard let parent = NSApp.mainWindow else { return 0 }
    parent.addChildWindow(window, ordered: .above)
    return 0
}

@_cdecl("TWHide")
public func TWHide(_ handle: Int64, _: Int64, _: Int64, _: Int64) -> Int64 {
    guard let window = textWindows[handle] else { return 0 }
    window.orderOut(nil)
    return 0
}

func getFocusedAccessibility(_ node: AnyObject) -> AnyObject? {
    let isAccessibilityFocusedSelector = #selector(NSAccessibilityProtocol.isAccessibilityFocused)
    let accessibilityChildrenSelector = #selector(NSAccessibilityProtocol.accessibilityChildren)

    let ref = _unsafeReferenceCast(node, to: NSAccessibilityProtocol.self)
    if ref.responds(to: isAccessibilityFocusedSelector) && ref.isAccessibilityFocused() {
        return node
    }

    if let children = ref.responds(to: accessibilityChildrenSelector)
        ? ref.accessibilityChildren() : nil
    {
        for child in children {
            if let result = getFocusedAccessibility(child as AnyObject) {
                return result
            }
        }
    }

    return nil
}

func getAccessibilityFrame(_ node: AnyObject) -> NSRect? {
    let accessibilityFrameSelector = #selector(
        NSAccessibilityProtocol.accessibilityFrame as (NSAccessibilityProtocol) -> () -> NSRect)
    guard node.responds(to: accessibilityFrameSelector) else { return nil }
    return _unsafeReferenceCast(node, to: NSAccessibilityProtocol.self).accessibilityFrame()
}

@_cdecl("TWSetResizeTarget")
public func SetResizeTarget(_ handle: Int64, _: Int64, _: Int64, _: Int64) -> Int64 {
    guard let window = textWindows[handle] else { return 0 }
    guard let parent = NSApp.mainWindow else { return 0 }
    guard let target = getFocusedAccessibility(parent) else { return 0 }
    window.resizeTarget = target
    return 0
}

@_cdecl("TWResize")
public func TWResize(_ handle: Int64, _: Int64, _: Int64, _: Int64) -> Int64 {
    guard let window = textWindows[handle] else { return 0 }
    guard let target = window.resizeTarget else { return 0 }
    guard let frame = getAccessibilityFrame(target) else { return 0 }
    window.setFrame(frame, display: true)
    return 0
}

@_cdecl("TWFocus")
public func TWFocus(_ handle: Int64, _: Int64, _: Int64, _: Int64) -> Int64 {
    guard let window = textWindows[handle] else { return 0 }
    window.makeKey()
    return 0
}

@_cdecl("TWGetByteLength")
public func TWGetByteLength(_ handle: Int64, _: Int64, _: Int64, _: Int64) -> Int64 {
    guard let window = textWindows[handle] else { return 0 }
    let string = window.textView.string as NSString
    return Int64(string.lengthOfBytes(using: String.Encoding.utf16LittleEndian.rawValue))
}

@_cdecl("TWGetBytes")
public func TWGetBytes(_ handle: Int64, buffer: UnsafeMutableRawPointer?, length: Int64, _: Int64)
    -> Int64
{
    guard let window = textWindows[handle] else { return 0 }
    let string = window.textView.string as NSString
    var usedLength: Int = 0
    string.getBytes(
        buffer,
        maxLength: Int(length),
        usedLength: &usedLength,
        encoding: String.Encoding.utf16LittleEndian.rawValue,
        options: [],
        range: NSRange(location: 0, length: Int(length)),
        remaining: nil
    )
    return Int64(usedLength)
}

@_cdecl("TWSetBytes")
public func TWSetBytes(_ handle: Int64, buffer: UnsafeRawPointer?, length: Int64, _: Int64) -> Int64
{
    guard let window = textWindows[handle] else { return 0 }
    if let buffer = buffer,
        let string = NSString(
            bytes: buffer,
            length: Int(length),
            encoding: String.Encoding.utf16LittleEndian.rawValue
        )
    {
        window.textView.string = string as String
    } else {
        window.textView.string = ""
    }
    return 0
}

@_cdecl("TWGetSelStart")
public func TWGetSelStart(_ handle: Int64, _: Int64, _: Int64, _: Int64) -> Int64 {
    guard let window = textWindows[handle] else { return 0 }
    return Int64(window.textView.selectedRange.location)
}

@_cdecl("TWSetSelStart")
public func TWSetSelStart(_ handle: Int64, value: Int64, _: Int64, _: Int64) -> Int64 {
    guard let window = textWindows[handle] else { return 0 }
    window.textView.selectedRange = NSRange(location: Int(value), length: 0)
    return 0
}

@_cdecl("TWGetFontSize")
public func TWGetFontSize(_ handle: Int64, _: Int64, _: Int64, _: Int64) -> Float64 {
    guard let window = textWindows[handle] else { return 0 }
    return Float64(window.textView.font?.pointSize ?? 0)
}

@_cdecl("TWSetFontSize")
public func TWSetFontSize(_ handle: Int64, value: Float64, _: Int64, _: Int64) -> Int64 {
    guard let window = textWindows[handle] else { return 0 }
    window.textView.font = NSFont.userFixedPitchFont(ofSize: CGFloat(value))
    return 0
}

@_cdecl("TWGetWordWrap")
public func TWGetWordWrap(_ handle: Int64, _: Int64, _: Int64, _: Int64) -> Int64 {
    guard let window = textWindows[handle] else { return 0 }
    return window.wordWrap ? 1 : 0
}

@_cdecl("TWSetWordWrap")
public func TWSetWordWrap(_ handle: Int64, value: Int64, _: Int64, _: Int64) -> Int64 {
    guard let window = textWindows[handle] else { return 0 }
    window.wordWrap = value != 0
    return 0
}
