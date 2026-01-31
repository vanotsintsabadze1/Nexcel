# Nexcel

> **Warning**
> This project is currently **experimental** and **unstable**. It is a work in progress and APIs may change without notice.

## Status & Limitations

Please note that this library is not feature-complete. It is an experimental utility and there is **no full commitment** to long-term maintenance or feature parity with established libraries. Use it for prototyping or simple use cases where its abstractions fit, but be aware of its limitations.

## Overview

Nexcel provides a high-level abstraction for building Excel documents in .NET.

Traditionally, generating Excel files using libraries like ClosedXML requires writing significant boilerplate code or implementing custom reflection logic to map objects to rows. Nexcel simplifies this process by handling the complexities of object mapping and styling, allowing developers to generate files on the go with minimal setup.

## Why Nexcel?

- **Simplified Usage:** Generates Excel files without the need for verbose low-level manipulation.
- **Abstraction:** Removes the need to manually handle reflection or iterate over collections to populate cells.
- **Focus:** Allows consumers to focus on their data rather than the mechanics of file generation.