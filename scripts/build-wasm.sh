#!/bin/sh
# Builds the rdocx wasm from ../rdocx (branch dopame) and copies it into the app:
# the JS glue and the .wasm into src/wasm/ (Vite bundles both).
set -eu

here="$(cd "$(dirname "$0")" && pwd)"
app="$(dirname "$here")"
rdocx="$(dirname "$app")/rdocx"

. "$HOME/.cargo/env"

cd "$rdocx"
cargo build -p rdocx-wasm --target wasm32-unknown-unknown --release
rm -rf target/pkg-web
wasm-bindgen target/wasm32-unknown-unknown/release/rdocx_wasm.wasm --out-dir target/pkg-web --target web

mkdir -p "$app/src/wasm"
cp target/pkg-web/rdocx_wasm.js target/pkg-web/rdocx_wasm.d.ts "$app/src/wasm/"
cp target/pkg-web/rdocx_wasm_bg.wasm "$app/src/wasm/rdocx_wasm_bg.wasm"
ls -l "$app/src/wasm/rdocx_wasm_bg.wasm"
