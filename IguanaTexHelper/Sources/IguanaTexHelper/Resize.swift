import AppKit

public typealias ResizeHandler = @convention(c) (_ formPtr: Int, Double, Double, Double, Double)
    -> Void

@_cdecl("MacMakeFormResizable")
public func MacMakeFormResizable(formPtr: Int, handler: @escaping ResizeHandler, _: Int64, _: Int64)
    -> Int64
{
    guard let window = NSApp.mainWindow else { return 0 }
    window.styleMask.insert(.resizable)

    NotificationCenter.default.addObserver(
        forName: NSWindow.didResizeNotification,
        object: window,
        queue: nil
    ) { _ in guard window.inLiveResize else { return }
        guard let screen = window.screen else { return }
        let frame = window.frame
        handler(
            formPtr,
            Double(frame.minX),
            Double(screen.frame.maxY - frame.maxY),
            Double(frame.width),
            Double(frame.height)
        )
    }

    return 0
}
