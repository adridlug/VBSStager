# VBSStager

> **Warning**
> This repository contains a multi-stage VBScript stager that downloads and executes a Windows binary. That behavior is commonly associated with malware delivery. Use this project **only** in an isolated lab environment for defensive research, malware analysis training, or controlled detection engineering.

## Overview

VBSStager is a small research project that demonstrates a staged VBScript delivery chain on Windows. A Python builder assembles several VBScript templates into a final launcher script.

At a high level, the generated script chain:

1. decodes an embedded second-stage VBScript,
2. writes it to disk,
3. downloads an executable from an HTTP server,
4. writes a third-stage VBScript to disk, and
5. launches the downloaded executable.

This makes the repository useful for:

- malware analysis labs,
- blue-team detection testing,
- YARA / Sigma / EDR rule validation,
- controlled adversary-emulation research.

## Repository contents

- `stager.py` — Python builder that reads the stage templates, base64-encodes embedded content, and produces `final.vbs`.
- `stage1_template.vbs` — first-stage template that decodes and writes the second stage.
- `stage2_template.vbs` — second-stage template that downloads a payload and writes the third stage.
- `stage3_template.vbs` — third-stage template that launches the downloaded executable.
- `final.vbs` — generated output script.
- `Rechnung.docm` / `Rechnung.7z` — sample files included in the repository.

## How it works

### Stage 1

`stage1_template.vbs` contains a placeholder for embedded base64 data. When the final script runs, it:

- creates `C:\tmp\` if needed,
- decodes the embedded stage 2 script,
- writes it as `C:\tmp\test.vbs`,
- executes it with `wscript`.

### Stage 2

`stage2_template.vbs`:

- downloads `testapp.exe` from `http://[IP]:[PORT]/testapp.exe`,
- saves it under `C:\tmp\`,
- decodes the embedded stage 3 script,
- writes it as `C:\tmp\test1.vbs`,
- executes the next script.

### Stage 3

`stage3_template.vbs` launches the downloaded executable:

- `C:\tmp\testapp.exe [IP]`

### Builder

`stager.py` is the glue that assembles everything:

- reads `stage3_template.vbs`, inserts the configured IP, and base64-encodes it,
- injects that result into `stage2_template.vbs`,
- injects IP and port values into stage 2,
- base64-encodes stage 2,
- splits the encoded string into VBScript-friendly chunks,
- injects the result into `stage1_template.vbs`,
- writes the final output to `final.vbs`.

## Requirements

- Windows test environment
- Python 3
- Windows Script Host / `wscript`
- A local HTTP server hosting the test payload

## Usage

### 1. Configure the builder

Edit the top of `stager.py`:

```python name=stager.py
IP = "192.168.178.67"
PORT = "8000"
```

Set these values to your lab listener or local web server.

### 2. Prepare a test payload

Host a Windows executable named `testapp.exe` from your HTTP server.

For example, the generated second stage expects:

- `http://<IP>:<PORT>/testapp.exe`

### 3. Generate the final script

Run:

```bash name=build.sh
python stager.py
```

This creates or updates:

- `final.vbs`

### 4. Execute only in an isolated lab

Run the generated VBScript inside a disposable virtual machine or analysis sandbox.

## Safe research guidance

If you use this repository for legitimate security work:

- prefer a non-functional benign payload for testing,
- use an offline or segmented lab network,
- avoid production hosts and personal systems,
- clearly label generated samples as research artifacts,
- rotate any IPs, domains, or infrastructure after tests.

## Detection ideas

Defenders can use this project to test detections for:

- `wscript.exe` spawning child scripts or executables,
- creation of files in `C:\tmp\` such as `test.vbs`, `test1.vbs`, and `testapp.exe`,
- use of `ADODB.Stream` and `Msxml2.DOMDocument` for base64 decoding and file writes,
- outbound HTTP requests from script interpreters,
- staged script-to-binary execution chains.

## Notes

- The current implementation uses hard-coded paths, filenames, IP address, and port.
- Error handling, cleanup, and safety guardrails are minimal.
- The repository appears intended for experimentation rather than production-quality tooling.

## Disclaimer

This project is provided for authorized research, education, and defensive testing only. Users are solely responsible for ensuring that their use complies with all applicable laws, regulations, and organizational policies. The author does not endorse or accept responsibility for any misuse, illegal activity, or damage resulting from the use of this project.
