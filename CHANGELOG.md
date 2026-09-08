# Changelog

## 1.3.0
- MIME `From:` and `To:` headers are authoritative for From/To names and addresses.
- Default fields: From Name, From Email, To Name, To Email, Subject.
- Removed CC Name, CC Email, Raw From Header, and Raw To Header from the field picker.
- Retained structured Outlook properties as fallbacks when MIME data cannot be read.
- XLSX export remains local and dependency-free.

## 1.2.0
- Removed slow Internet-header calls from the default export path.
- Added bounded fallback behaviour for messages with incomplete structured recipient data.

## 1.1.0
- Added selectable export fields.
- Added direct XLSX export.
- Added separate From and Sender/Sent-by fields.
- Removed sender-summary export.

## 1.0.0
- Initial New Outlook multi-select exporter.
- CSV detail and sender-summary exports.
