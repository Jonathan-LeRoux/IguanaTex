SWIFT := swift
SWIFT_FORMAT := swift-format

SWIFT_BUILD_FLAGS := -c release --arch arm64 --arch x86_64
SWIFT_FORMAT_FLAGS := --configuration .swift-format --in-place

SWIFT_BIN_PATH := $(shell $(SWIFT) build $(SWIFT_BUILD_FLAGS) --show-bin-path)
ADDIN_DIR := /Library/Application Support/Microsoft/Office365/User Content.localized/Add-Ins.localized/

.PHONY: help all format install show-bin-path clean

all: ## Build the library
	$(SWIFT) build $(SWIFT_BUILD_FLAGS)

install: ## Install the library (requires root privileges)
	install -m755 '$(SWIFT_BIN_PATH)/libIguanaTexHelper.dylib' '$(DESTDIR)$(ADDIN_DIR)'

show-bin-path: ## Print the binary output path
	@echo '$(SWIFT_BIN_PATH)'

clean: ## Delete build artifacts
	$(SWIFT) package clean
	rm -rf .build

format: ## Format source files using `swift-format`
	$(SWIFT_FORMAT) format $(SWIFT_FORMAT_FLAGS) --recursive Sources

help: ## Print this help message
	@grep -E '^[a-zA-Z._-]+:.*?## .*$$' $(MAKEFILE_LIST) | sort | awk 'BEGIN {FS = ":.*?## "}; {printf "\033[36m%-30s\033[0m %s\n", $$1, $$2}'
