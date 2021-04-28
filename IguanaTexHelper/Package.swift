// swift-tools-version:5.3

import PackageDescription

let package = Package(
    name: "IguanaTexHelper",
    products: [
        .library(
            name: "IguanaTexHelper",
            type: .dynamic,
            targets: ["IguanaTexHelper"]),
    ],
    dependencies: [
    ],
    targets: [
        .target(
            name: "IguanaTexHelper",
            dependencies: []),
    ]
)
