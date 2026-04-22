# Product.wxs — Required Changes for Faster-Whisper Architecture

Apply each change block below to `WorshipHelperSetup/Product.wxs`.

---

## Change 1 — Replace `ComponentGroupRef` (line ~14)

```xml
<!-- BEFORE -->
<ComponentGroupRef Id="VoskModelFiles" />

<!-- AFTER -->
<ComponentGroupRef Id="SttServerFiles" />
<ComponentRef Id="VersesDb_Component" />
<ComponentRef Id="StartSttServer_Component" />
```

---

## Change 2 — Replace `VOSK_MODEL_DIR` with `STT_SERVER_DIR` (line ~44)

```xml
<!-- BEFORE -->
<Directory Id="DATA_DIR" Name="data">
    <Directory Id="VOSK_MODEL_DIR" Name="vosk-model" />
    <Directory Id="TEMPLATES_DIR"  Name="Templates" />
    <Directory Id="BIBLES_DIR"     Name="Bibles" />
</Directory>

<!-- AFTER -->
<Directory Id="DATA_DIR" Name="data">
    <Directory Id="TEMPLATES_DIR" Name="Templates" />
    <Directory Id="BIBLES_DIR"    Name="Bibles" />
</Directory>
<Directory Id="STT_SERVER_DIR" Name="stt-server" />
```

---

## Change 3 — Remove Vosk DLL components, add SQLite + Json.NET

Remove these five Components (approximately lines 114–148):
- `Vosk_dll_Component`
- `libvosk_dll_Component`
- `libgcc_dll_Component`
- `libstdcpp_dll_Component`
- `libwinpthread_dll_Component`

Add in their place, inside `<ComponentGroup Id="ProductComponents">`:

```xml
<Component Id="SQLite_dll_Component">
    <File Id="SQLite_dll" KeyPath="yes"
          Name="System.Data.SQLite.dll"
          Source="$(var.WorshipHelperVSTO.TargetDir)" />
</Component>
<Component Id="SQLiteInterop_dll_Component">
    <File Id="SQLiteInterop_dll" KeyPath="yes"
          Name="SQLite.Interop.dll"
          Source="$(var.WorshipHelperVSTO.TargetDir)" />
</Component>
<Component Id="NewtonsoftJson_dll_Component">
    <File Id="NewtonsoftJson_dll" KeyPath="yes"
          Name="Newtonsoft.Json.dll"
          Source="$(var.WorshipHelperVSTO.TargetDir)" />
</Component>
```

---

## Change 4 — Add `verses.sqlite` component under `DATA_DIR`

```xml
<DirectoryRef Id="DATA_DIR">
    <Component Id="VersesDb_Component">
        <File Id="verses_sqlite" KeyPath="yes"
              Name="verses.sqlite"
              Source="$(var.WorshipHelperVSTO.TargetDir)data\verses.sqlite" />
    </Component>
</DirectoryRef>
```

---

## Change 5 — Add start script component under `STT_SERVER_DIR`

```xml
<DirectoryRef Id="STT_SERVER_DIR">
    <Component Id="StartSttServer_Component">
        <File Id="start_stt_server_bat" KeyPath="yes"
              Name="start-stt-server.bat"
              Source="$(var.WorshipHelperVSTO.TargetDir)stt-server\start-stt-server.bat" />
    </Component>
</DirectoryRef>
```
