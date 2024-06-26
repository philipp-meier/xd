# xd

[![MIT License](https://img.shields.io/badge/license-MIT-green.svg)](https://github.com/philipp-meier/xd/blob/main/LICENSE)

xd is a simple Excel (.**x**lsx) **d**iff tool for texts.  

## Usage
```bash
# Run main.go with parameters
go run main.go -f1 ./data/InputA.xlsx -f2 ./data/InputB.xlsx

# Build, "publish" and execute
go build
# env GOOS=linux GOARCH=amd64 go build -o xd_linux_amd64
# env GOOS=windows GOARCH=amd64 go build -o xd_windows_amd64.exe
mv xd /usr/bin
xd -f1 ./data/InputA.xlsx -f2 ./data/InputB.xlsx

# Print help
xd -h

# Measure execution time
time xd -f1 ./data/InputA.xlsx -f2 ./data/InputB.xlsx
```

## Sample output
```bash
CAUTION: File ./data/InputB.xlsx has no sheet called "Table21"
Sheet1
- A1: Hello <> Hello1
- B1: World <> World2
Sheet2
- E12: Example <> Test
```

## Limitations
This tool currently requires both Excel files to have the same structure.  
Adding a new line in the middle of a sheet would therefore lead to many "differences".  
