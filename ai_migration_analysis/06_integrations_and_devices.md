# 06 Integrations And Devices

## Tally Integration

Sources:

- `../LegacyCodebase/ModTallyInterface.bas`
- `../LegacyCodebase/modProcess.bas`
- `../LegacyCodebase/modXml.bas`
- `../LegacyCodebase/ETLInterfaceForm.frm`
- `../LegacyCodebase/INDEXFILES/rtslink.dll`

Technology:

- RTSlink COM/DLL
- XML payloads
- likely local Tally instance endpoint in `strURI`

Confirmed functions:

- `ConnectToTally`
- `SendData`
- `GetData`
- `CreatePrimaryGroups`
- `CreateGroups`
- `CreateAccounts`
- `CreateUOM`
- `CreateStockGroup`
- `CreateStockItem`
- `CreateAccountingVoucher`
- `CreateInventoryVoucher`

Behavior:

- XML payload is string-built.
- `SendData` replaces `&` with `&amp;`.
- Tally response is parsed through substring searches for `<ERRORS>` and `<LINEERROR>`.
- Errors are appended to a `RichTextBox`.

Migration recommendation:

- Replace string slicing with XML parser.
- Use typed integration events and retries.
- Persist every outbound payload and inbound response.
- Make Tally sync asynchronous through background jobs.

## HHT / Handheld Terminal

Sources:

- `../LegacyCodebase/M_FRM_MIBS.frm`
- `../LegacyCodebase/HHT/*.TXT`
- `../LegacyCodebase/INDEXFILES/HHT`

Technology:

- `MSCOMM32.OCX`
- serial port communication
- flat files

Known HHT files:

- `ACT_PAY.TXT`
- `ARCOLGOD.TXT`
- `EMPTIES.TXT`
- `EMPTY_ISSUE.TXT`
- `INL_STR.TXT`
- `INV_STR.TXT`
- `LOAD_SHEET.TXT`
- `LOCATION.TXT`
- `LOG.TXT`
- `MARKET.TXT`
- `PARTY.TXT`
- `SALESMAN.TXT`
- `STK_TRN.TXT`
- `STK_TRN_AMT.TXT`
- `STOCK.TXT`
- `VERSION.TXT`

Observed behavior:

- App receives HHT identity through serial input.
- Validates identity against `HHT` table.
- Writes outbound files.
- Reads inbound `INL_STR.TXT` and `INV_STR.TXT`.
- Updates core transaction and inventory tables.

Migration recommendation:

- Short term: preserve import/export compatibility for existing HHT files.
- Long term: replace with offline-first mobile/PWA sync.
- Model sync as explicit sessions with validation, audit, conflict handling, and replayable events.

## SMS Integration

Sources:

- `../LegacyCodebase/modInnComm.bas`
- `../LegacyCodebase/modSMSSend.bas`
- `../LegacyCodebase/modMessageTemplate.bas`
- `../LegacyCodebase/frmSMS.frm`
- logs in `INDEXFILES/LOG`

Technology:

- serial modem via `MSCOMM32.OCX`
- AT commands
- plain text logs

Behavior:

- `SendMessage` opens COM port, splits message, writes AT command payload, waits for response.
- `ReceiveMessage` queries SMS list.
- `DeleteMessage` deletes modem-stored messages.
- `CommSettings` reads settings from setup files.

Migration recommendation:

- Replace modem with provider API if SMS is still required.
- Keep logs/audit.
- Store message templates in DB.

## Barcode / Label Printing

Dependencies:

- BarTender 8.01
- `bartend.exe`
- `frmBarCode.frm`
- `frmNewRateTemplate.frm`

Observed source:

- New rate/batch templates can insert/update `BRD`.
- Barcode/template command handlers exist.

Migration recommendation:

- Treat barcode labels as generated PDF/ZPL/template outputs.
- Confirm actual printer hardware before implementation.

## Reporting

Report assets:

- `.RP1`
- `.RP5`
- `.RRW`
- `RRW32.OCX`
- `RRI.DBF`
- `pri.exe`

Report domains:

- stock
- sales
- route
- market
- retailer
- salesman
- SKU
- finance
- aging
- VAT
- assets
- productivity

Migration recommendation:

- Build reporting views in Postgres.
- Recreate high-value reports first.
- Do not port report binaries directly.

## Licensing / Dongle

Sources:

- `CL32.DLL`
- `cl_login`, `cl_logout`, `cl_lock_ok` declarations in `EFF_LIB.bas`
- setup files under Windows system path (`EFFMCG4.txt`)

Migration recommendation:

- Replace with modern license/tenant configuration.
- Preserve no dongle-specific behavior unless contractually required.
