import AppKit
import Carbon
import InterposeKit

public typealias RawCopyPasteHandler = @convention(c) (
    _ formPtr: Int, _ keyCode: Int64, _ modifierFlags: Int64
) -> Void

public typealias RawAcceleratorHandler = @convention(c) (
    _ formPtr: Int, _ asciiCode: Int64
) -> Void

public typealias CopyPasteHandler = (
    _ keyCode: Int64, _ modifierFlags: Int64
) -> Void

public typealias AcceleratorHandler = (
    _ asciiCode: Int64
) -> Void

@_cdecl("MacEnableCopyPaste")
public func MacEnableCopyPaste(
    formPtr: Int,
    handler: @escaping RawCopyPasteHandler,
    _: Int64,
    _: Int64
) -> Int64 {
    guard let window = NSApp.mainWindow else { return 0 }
    guard let _ = interposer else { return 0 }
    window.copyPasteHandler = { (keyCode, modifierFlags) in
        handler(formPtr, keyCode, modifierFlags)
    }
    return 0
}

@_cdecl("MacEnableAccelerators")
public func MacEnableAccelerators(
    formPtr: Int,
    handler: @escaping RawAcceleratorHandler,
    _: Int64,
    _: Int64
) -> Int64 {
    guard let window = NSApp.mainWindow else { return 0 }
    guard let _ = interposer else { return 0 }
    window.acceleratorHandler = { asciiCode in handler(formPtr, asciiCode) }
    return 0
}

let interposer = try? Interpose(NSApplication.self) {
    try $0.hook(
        #selector(NSWindow.sendEvent(_:)),
        methodSignature: (@convention(c) (AnyObject, Selector, NSEvent) -> Void).self,
        hookSignature: (@convention(block) (AnyObject, NSEvent) -> Void).self
    ) { store in
        { `self`, event in
            if let newEvent = handleEvent(event) {
                store.original(`self`, store.selector, newEvent)
            }
        }
    }
}

func handleEvent(_ event: NSEvent) -> NSEvent? {
    guard
        event.type == .keyDown,
        let targetWindow = event.window,
        let mainWindow = NSApp.mainWindow
    else {
        return event
    }

    let keyModifierFlags = event.modifierFlags.intersection(.deviceIndependentFlagsMask)

    // Custom handler for copy, paste, undo, redo...
    if let copyPasteHandler = targetWindow.copyPasteHandler {
        switch (keyModifierFlags, Int(event.keyCode)) {
        case (.command, kVK_ANSI_C),
            (.command, kVK_ANSI_V),
            (.command, kVK_ANSI_X),
            (.command, kVK_ANSI_A),
            (.command, kVK_ANSI_Z),
            ([.command, .shift], kVK_ANSI_Z):
            copyPasteHandler(Int64(event.keyCode), Int64(keyModifierFlags.rawValue))
            return nil
        default:
            break
        }
    }

    // Disable command+* keys if accelerators are enabled
    if targetWindow.acceleratorHandler != nil && keyModifierFlags == [.command] {
        return nil
    }

    // Map control+command+* to accelerators
    if keyModifierFlags == [.control, .command],
        let char = event.charactersIgnoringModifiers?.first,
        char.isLetter && char.isASCII,
        let asciiCode = char.asciiValue,
        let acceleratorHandler = mainWindow.acceleratorHandler
    {
        acceleratorHandler(Int64(asciiCode))
        return nil
    }

    return event
}

struct Keys {
    static var copyPasteHandler: UInt8 = 0
    static var acceleratorHandler: UInt8 = 0
}

extension NSWindow {
    fileprivate var copyPasteHandler: CopyPasteHandler? {
        get { objc_getAssociatedObject(self, &Keys.copyPasteHandler) as? CopyPasteHandler }
        set {
            objc_setAssociatedObject(
                self,
                &Keys.copyPasteHandler,
                newValue,
                .OBJC_ASSOCIATION_RETAIN_NONATOMIC
            )
        }
    }
    fileprivate var acceleratorHandler: AcceleratorHandler? {
        get { objc_getAssociatedObject(self, &Keys.acceleratorHandler) as? AcceleratorHandler }
        set {
            objc_setAssociatedObject(
                self,
                &Keys.acceleratorHandler,
                newValue,
                .OBJC_ASSOCIATION_RETAIN_NONATOMIC
            )
        }
    }
}
