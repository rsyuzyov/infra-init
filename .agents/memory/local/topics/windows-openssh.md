# Windows OpenSSH (sshd для Windows)

## Факты

- OpenSSH-сервер на Windows различает обычных пользователей и админов **по группе** `BUILTIN\Administrators` (SID `S-1-5-32-544`, `Match Group administrators` в дефолтном `sshd_config`). Локализация имени группы не важна — `Match` смотрит на SID.

## authorized_keys: два места

| Кто | Файл |
|---|---|
| Обычный пользователь | `C:\Users\<user>\.ssh\authorized_keys` |
| Член группы `Administrators` | `C:\ProgramData\ssh\administrators_authorized_keys` (общий файл) |

⚠️ Для админа файл в `~/.ssh/authorized_keys` **игнорируется** — sshd идёт сразу в общий админский файл.

## Как sshd выбирает пользователя сессии

- Имя юзера приходит от клиента: `ssh user@host`.
- sshd резолвит имя, проверяет членство в `administrators`, выбирает соответствующий файл authorized_keys.
- При совпадении ключа сессия запускается **под тем юзером, чьё имя передал клиент**, не "под владельцем ключа".
- Один ключ в `administrators_authorized_keys` пустит под любым админом, чьё имя передал клиент.

## ACL — критично

sshd жёстко валидирует права на файлы ключей. При несоответствии — `bad permissions` в журнале sshd и тихий отказ в логине.

| Файл | Владелец | Доступ только у | Прочие требования |
|---|---|---|---|
| `C:\Users\<user>\.ssh\authorized_keys` | сам user | user + SYSTEM | `icacls /inheritance:r`, убрать Administrators |
| `C:\ProgramData\ssh\administrators_authorized_keys` | Administrators или SYSTEM | Administrators + SYSTEM | `icacls /inheritance:r` |

### Команды

```powershell
# User-файл
icacls $f /inheritance:r
icacls $f /grant "$user`:F" "SYSTEM:F"

# Admin-файл
icacls $f /inheritance:r
icacls $f /grant "BUILTIN\Administrators:F" "SYSTEM:F"
```

## Грабли

- ⚠️ Положить ключ для админа в `~/.ssh/authorized_keys` — типовая ошибка. Не сработает, нужен общий `administrators_authorized_keys`.
- ⚠️ Оставить наследование на админском файле от `C:\ProgramData\ssh` — sshd начнёт ругаться на права. `/inheritance:r` обязателен.
- ⚠️ Не убрать `Administrators` из ACL user-файла → sshd тоже считает права слишком широкими.
- ⚠️ Опции в строке ключа (`from="..."`, `command="..."`) — стандартный способ ограничить применение ключа конкретным источником/командой. Особенно полезно при общем админском файле.
