# RadEditSync Bridge Protocol (v1.0)

This protocol defines how the local client script and the Citrix worker exchange transfer jobs through the mounted `RadEditSync` directory.

## Goals
- Keep local `SendCopyData` / `CopyDataHandler` unchanged.
- Cross the Citrix boundary using files only.
- Make each transfer auditable and retry-safe.

## Directory Layout

```text
RadEditSync/
  queue/
    <jobId>/
      meta.json
      report.rtf
      textefinal.rtf      (optional, signature/footer RTF to append)
      textesRapport.txt   (optional, fallback text dictionary)
      ready.flag
      done.json          (written by Citrix worker on success)
      error.json         (written by Citrix worker on failure)
      processing.lock    (optional, worker-owned)
  archive/               (optional retention target)
```

`jobId` format recommendation:
- `job_YYYYMMDD_HHMMSS_<6digits>`

## Job Lifecycle
1. Local side creates `queue/<jobId>/`.
2. Local side writes `meta.json` and `report.rtf`.
3. Local side writes `ready.flag` last (publish step).
4. Citrix worker picks a ready job and may create `processing.lock`.
5. Citrix worker executes RadImage transfer.
6. Citrix worker writes exactly one terminal file:
- success: `done.json`
- failure: `error.json`
7. Local side reads terminal file, then performs local cleanup.

## Atomicity Rules
- `ready.flag` must be created only after `meta.json` and `report.rtf` are fully written.
- Worker must ignore jobs without `ready.flag`.
- Worker should skip jobs already containing `done.json` or `error.json`.

## `meta.json` (required keys)

All values are strings in protocol v1.0.

```json
{
  "protocolVersion": "1.0",
  "jobId": "job_20260227_150501_123456",
  "createdAtUtc": "2026-02-27T15:05:01Z",
  "mode": "ToutCitrix",
  "reqnb": "RA202600000001",
  "patdos": "1234567",
  "patnom": "DOE, JOHN",
  "proc": "CT THORAX",
  "studydate": "2026-02-27",
  "stripHiddenMarkers": "true",
  "keepfont": "false",
  "signatureFile": "textefinal.rtf",
  "attv": "false",
  "casexterne": "false"
}
```

## Optional Job Artifacts
- `textefinal.rtf`: if present in the job folder, Citrix uses it first for end/signature insertion.
- `textesRapport.txt`: optional fallback text source for signature block fields.

## `done.json` (terminal success)

```json
{
  "protocolVersion": "1.0",
  "jobId": "job_20260227_150501_123456",
  "completedAtUtc": "2026-02-27T15:05:29Z",
  "status": "done",
  "worker": "citrix"
}
```

## `error.json` (terminal failure)

```json
{
  "protocolVersion": "1.0",
  "jobId": "job_20260227_150501_123456",
  "completedAtUtc": "2026-02-27T15:05:29Z",
  "status": "error",
  "worker": "citrix",
  "errorCode": "RADIMAGE_TIMEOUT",
  "message": "RadImage window not active within timeout"
}
```

## Timeouts and Retry
- Local wait timeout recommendation: 60 seconds (configurable).
- If timeout occurs with no terminal file, local side may re-trigger Citrix hotkey and continue waiting.
- Worker should be idempotent per `jobId`; never process the same terminal job twice.

## Compatibility
- Any future breaking change must increment `protocolVersion`.
- Unknown keys must be ignored by both sides.
