#!/usr/bin/env python3
"""Check the rdocx files bundled with HelloWord."""

from pathlib import Path


def main() -> None:
    bundle = Path(__file__).resolve().parents[1] / "src" / "wasm"
    module = bundle / "rdocx_wasm_bg.wasm"
    for path in (module, bundle / "rdocx_wasm.js", bundle / "rdocx_wasm.d.ts"):
        if not path.is_file() or not path.stat().st_size:
            raise SystemExit(f"Missing or empty rdocx bundle file: {path}")

    with module.open("rb") as wasm:
        if wasm.read(8) != b"\x00asm\x01\x00\x00\x00":
            raise SystemExit(f"Invalid WebAssembly v1 header: {module}")

    print("The rdocx WebAssembly module and JavaScript bindings are present.")


if __name__ == "__main__":
    main()
