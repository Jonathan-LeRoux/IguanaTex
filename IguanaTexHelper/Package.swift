// swift-tools-version:5.3

import PackageDescription

let package = Package(
    name: "IguanaTexHelper",
    platforms: [.macOS(.v10_13)],
    products: [
        .library(
            name: "IguanaTexHelper",
            type: .dynamic,
            targets: ["IguanaTexHelper"]),
    ],
    dependencies: [
        .package(url: "https://github.com/steipete/InterposeKit.git", from: "0.0.2"),
    ],
    targets: [
        .target(
            name: "IguanaTexHelper",
            dependencies: ["InterposeKit"]),
    ]
)
