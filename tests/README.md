# Проверки ProfileMigrator

По требованию владельца проекта тесты с изменением реального реестра на его ПК запрещены, включая временные разделы HKCU. Не запускать для проверки ни `CreateProfile`, ни реальную регистрацию SID, ни сам GUI/EXE с операциями миграции. Не повторять вызовы переименования ключей реестра.

Все следующие проверки выполняются в отдельных процессах Windows PowerShell 5.1:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File tests/TargetFoldersDialog.Tests.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File tests/CreateProfileCall.Tests.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File tests/ProfileRegistryWriter.Tests.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File tests/ExistingProfileAdoption.Tests.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File tests/ProfileRegistration.Tests.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File tests/TargetFolders.Tests.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File tests/SelectedUserOperations.Tests.ps1
```

- `ProfileRegistryWriter`: выполняет реальную логику заполнения записи с хранилищем в памяти. Проверяет значения (включая FullProfile DWORD 1), добавление только отсутствующего FullProfile в существующую запись на шаге 2, защиту существующего SID, отмену, отказ каждой записи и откат только собственной записи. Объявления нативного адаптера компилируются, но адаптер не вызывается.
- `ExistingProfileAdoption`: использует временные файлы и имитацию реестра, AD, CIM и изменения ACL. Проверяет сохранение пользовательских файлов и NTUSER.DAT, дополнение из Default, повторный запуск, отсутствие дубликатов, конфликты и ошибки. Метаданные ссылок имитируются, junction на ПК не создаются.
- `CreateProfileCall`: проверяет фактическую обёртку вызова CreateProfile на управляемой имитации: буфер MAX_PATH (260), SID, логин, путь и передачу HRESULT. Настоящий userenv.dll не вызывается.
- `TargetFoldersDialog`: проверяет только окно выбора папок, отмену, добавление пользовательских путей, валидацию, сохранение и повторную загрузку из временного JSON. Рабочая конфигурация и операции миграции не запускаются.
- Остальные проверки используют временные данные; реальные изменения профилей, реестра и ACL подменены. Robocopy копирует только локальные тестовые файлы.

Эти проверки не подтверждают вход пользователя в Windows/RDS. Проверку настоящего ProfileList и входа выполнять отдельно на тестовой VM со снимком, без работающих сеансов тестового пользователя. Для существующей папки ожидается тот же путь в ProfileImagePath, сохраненный NTUSER.DAT и отсутствие папок с суффиксами. При повторном шаге 2 используется существующая регистрация.

Регистрация папки не заменяет имеющиеся файлы. Последующая обычная миграция TargetFolders сохраняет прежнюю логику Robocopy и может обновлять отличающиеся файлы источником.
