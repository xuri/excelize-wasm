DIST = ./dist
SRC = ./src
WASM = ${DIST}/excelize.wasm

build:
	GOOS=js GOARCH=wasm go test -exec="${GOROOT}/misc/wasm/go_js_wasm_exec" ${SRC}
	npm install
	node ./node_modules/.bin/rollup -c
	GOOS=js GOARCH=wasm CGO_ENABLED=0 go build -v -a -ldflags="-w -s" \
		-gcflags=-trimpath=${GOPATH} \
		-asmflags=-trimpath=${GOPATH} \
		-o ${WASM} ${SRC}/main.go
	gzip -f --best ${WASM}
	cp excelize-wasm.svg ${DIST}
	cp chart.png ${DIST}
	cp LICENSE ${DIST}
	cp README.md ${DIST}
	cp ${SRC}/package.json ${DIST}
	cp ${SRC}/index.d.ts ${DIST}
