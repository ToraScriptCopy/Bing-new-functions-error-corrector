"""
Interactive Windows conflict scanner and fixer
- Language selection at start (10 languages)
- Auto-install psutil if missing (with automatic restart)
- Choose 1+ report formats from 10 options (.txt, .json, .csv, .xml, .html, .md, .log, .yml, .ini, .pdf)
- Localized prompts in chosen language
"""

import subprocess
import sys
import platform
import os
import json
from datetime import datetime
from typing import Tuple, List, Dict, Any

# ---------------------------
# Languages (native names)
# ---------------------------
LANGS = {
    "ru": "Русский",
    "en": "English",
    "es": "Español",
    "pt": "Português",
    "tr": "Türkçe",
    "de": "Deutsch",
    "fr": "Français",
    "it": "Italiano",
    "zh": "中文 (简体)",
    "ja": "日本語"
}

# ---------------------------
# Full translations for UI strings (all 10 languages)
# ---------------------------
TRANSLATIONS: Dict[str, Dict[str, str]] = {
    "ru": {
        "choose_lang_header": "Выберите язык / Select language:",
        "enter_number": "Введите номер (1-10):",
        "invalid_choice": "Неверный выбор. Попробуйте ещё раз.",
        "installing_psutil": "Модуль psutil не найден — выполняется автоматическая установка...",
        "psutil_installed": "psutil установлен. Перезапуск скрипта...",
        "psutil_failed": "Не удалось установить psutil: {err}",
        "scanning": "Сканирование системы на предмет потенциальных конфликтов...",
        "results_short": "Результаты (кратко):",
        "processes": "Процессы",
        "startup": "Win32 StartupCommand",
        "hkcu": "HKCU Run",
        "services": "Службы",
        "interactive_prompt": "Хотите пройти интерактивное разрешение найденных проблем?",
        "yes": "Да",
        "no": "Нет — только отчёт",
        "choose_report_formats": "Выберите форматы отчёта (через запятую, номера):",
        "formats_list": "1:.txt 2:.json 3:.csv 4:.xml 5:.html 6:.md 7:.log 8:.yml 9:.ini 10:.pdf",
        "generating_reports": "Генерация отчётов в выбранных форматах...",
        "done": "Готово.",
        "json_file": "JSON файл",
        "txt_file": "Текстовый файл",
        "press_enter": "Нажмите Enter для завершения...",
        "auto_install_reportlab": "Для создания PDF требуется пакет reportlab — выполняется установка...",
        "reportlab_failed": "Не удалось установить reportlab: {err}",
        "pdf_created": "PDF создан: {path}",
        "action_prompt": "Действие?",
        "kill": "Завершить",
        "skip": "Пропустить",
        "alternatives": "Альтернативы",
        "check_hkcu": "Проверить HKCU",
        "remove": "Удалить",
        "remove_prompt": "Удалить запись?",
        "stop_disable": "Остановить и отключить",
        "skip_label": "Пропустить",
        "failed_save": "Ошибка при сохранении {ext}: {err}"
    },
    "en": {
        "choose_lang_header": "Choose language / Выбор языка:",
        "enter_number": "Enter number (1-10):",
        "invalid_choice": "Invalid choice. Try again.",
        "installing_psutil": "psutil not found — installing automatically...",
        "psutil_installed": "psutil installed. Restarting script...",
        "psutil_failed": "Failed to install psutil: {err}",
        "scanning": "Scanning system for potential conflicts...",
        "results_short": "Results (short):",
        "processes": "Processes",
        "startup": "Win32 StartupCommand",
        "hkcu": "HKCU Run",
        "services": "Services",
        "interactive_prompt": "Do you want to run interactive remediation?",
        "yes": "Yes",
        "no": "No — report only",
        "choose_report_formats": "Choose report formats (comma-separated indices):",
        "formats_list": "1:.txt 2:.json 3:.csv 4:.xml 5:.html 6:.md 7:.log 8:.yml 9:.ini 10:.pdf",
        "generating_reports": "Generating reports in selected formats...",
        "done": "Done.",
        "json_file": "JSON file",
        "txt_file": "Text file",
        "press_enter": "Press Enter to exit...",
        "auto_install_reportlab": "PDF generation requires reportlab — installing...",
        "reportlab_failed": "Failed to install reportlab: {err}",
        "pdf_created": "PDF created: {path}",
        "action_prompt": "Action?",
        "kill": "Kill",
        "skip": "Skip",
        "alternatives": "Alternatives",
        "check_hkcu": "Check HKCU",
        "remove": "Remove",
        "remove_prompt": "Remove entry?",
        "stop_disable": "Stop+Disable",
        "skip_label": "Skip",
        "failed_save": "Failed to save {ext}: {err}"
    },
    "es": {
        "choose_lang_header": "Seleccione idioma / Select language:",
        "enter_number": "Ingrese número (1-10):",
        "invalid_choice": "Elección inválida. Intente de nuevo.",
        "installing_psutil": "psutil no encontrado — instalando automáticamente...",
        "psutil_installed": "psutil instalado. Reiniciando script...",
        "psutil_failed": "No se pudo instalar psutil: {err}",
        "scanning": "Escaneando el sistema en busca de conflictos potenciales...",
        "results_short": "Resultados (resumen):",
        "processes": "Procesos",
        "startup": "Win32 StartupCommand",
        "hkcu": "HKCU Run",
        "services": "Servicios",
        "interactive_prompt": "¿Desea ejecutar la remediación interactiva?",
        "yes": "Sí",
        "no": "No — sólo informe",
        "choose_report_formats": "Elija formatos de informe (separados por comas, índices):",
        "formats_list": "1:.txt 2:.json 3:.csv 4:.xml 5:.html 6:.md 7:.log 8:.yml 9:.ini 10:.pdf",
        "generating_reports": "Generando informes en los formatos seleccionados...",
        "done": "Hecho.",
        "json_file": "Archivo JSON",
        "txt_file": "Archivo de texto",
        "press_enter": "Presione Enter para salir...",
        "auto_install_reportlab": "Para generar PDF se necesita reportlab — instalando...",
        "reportlab_failed": "No se pudo instalar reportlab: {err}",
        "pdf_created": "PDF creado: {path}",
        "action_prompt": "Acción?",
        "kill": "Finalizar",
        "skip": "Omitir",
        "alternatives": "Alternativas",
        "check_hkcu": "Comprobar HKCU",
        "remove": "Eliminar",
        "remove_prompt": "¿Eliminar entrada?",
        "stop_disable": "Detener+Deshabilitar",
        "skip_label": "Omitir",
        "failed_save": "Error al guardar {ext}: {err}"
    },
    "pt": {
        "choose_lang_header": "Escolha o idioma / Select language:",
        "enter_number": "Digite o número (1-10):",
        "invalid_choice": "Escolha inválida. Tente novamente.",
        "installing_psutil": "psutil não encontrado — instalando automaticamente...",
        "psutil_installed": "psutil instalado. Reiniciando o script...",
        "psutil_failed": "Falha ao instalar psutil: {err}",
        "scanning": "Verificando o sistema em busca de possíveis conflitos...",
        "results_short": "Resultados (resumo):",
        "processes": "Processos",
        "startup": "Win32 StartupCommand",
        "hkcu": "HKCU Run",
        "services": "Serviços",
        "interactive_prompt": "Deseja executar a correção interativa?",
        "yes": "Sim",
        "no": "Não — apenas relatório",
        "choose_report_formats": "Escolha formatos de relatório (separados por vírgula, índices):",
        "formats_list": "1:.txt 2:.json 3:.csv 4:.xml 5:.html 6:.md 7:.log 8:.yml 9:.ini 10:.pdf",
        "generating_reports": "Gerando relatórios nos formatos selecionados...",
        "done": "Concluído.",
        "json_file": "Arquivo JSON",
        "txt_file": "Arquivo de texto",
        "press_enter": "Pressione Enter para sair...",
        "auto_install_reportlab": "Para gerar PDF é necessário reportlab — instalando...",
        "reportlab_failed": "Falha ao instalar reportlab: {err}",
        "pdf_created": "PDF criado: {path}",
        "action_prompt": "Ação?",
        "kill": "Finalizar",
        "skip": "Pular",
        "alternatives": "Alternativas",
        "check_hkcu": "Checar HKCU",
        "remove": "Remover",
        "remove_prompt": "Remover entrada?",
        "stop_disable": "Parar+Desabilitar",
        "skip_label": "Pular",
        "failed_save": "Falha ao salvar {ext}: {err}"
    },
    "tr": {
        "choose_lang_header": "Dil seçin / Select language:",
        "enter_number": "Sayı girin (1-10):",
        "invalid_choice": "Geçersiz seçim. Tekrar deneyin.",
        "installing_psutil": "psutil bulunamadı — otomatik yükleme yapılıyor...",
        "psutil_installed": "psutil yüklendi. Betik yeniden başlatılıyor...",
        "psutil_failed": "psutil yüklenemedi: {err}",
        "scanning": "Olası çakışmalar için sistem taranıyor...",
        "results_short": "Sonuçlar (kısa):",
        "processes": "İşlemler",
        "startup": "Win32 StartupCommand",
        "hkcu": "HKCU Run",
        "services": "Hizmetler",
        "interactive_prompt": "Etkileşimli düzeltme yapmak istiyor musunuz?",
        "yes": "Evet",
        "no": "Hayır — sadece rapor",
        "choose_report_formats": "Rapor formatlarını seçin (virgülle ayırın, numaralar):",
        "formats_list": "1:.txt 2:.json 3:.csv 4:.xml 5:.html 6:.md 7:.log 8:.yml 9:.ini 10:.pdf",
        "generating_reports": "Seçilen formatlarda raporlar oluşturuluyor...",
        "done": "Tamamlandı.",
        "json_file": "JSON dosyası",
        "txt_file": "Metin dosyası",
        "press_enter": "Çıkmak için Enter'a basın...",
        "auto_install_reportlab": "PDF oluşturmak için reportlab gerekiyor — yükleniyor...",
        "reportlab_failed": "reportlab yüklenemedi: {err}",
        "pdf_created": "PDF oluşturuldu: {path}",
        "action_prompt": "Eylem?",
        "kill": "Durdur",
        "skip": "Atla",
        "alternatives": "Alternatifler",
        "check_hkcu": "HKCU'yi kontrol et",
        "remove": "Kaldır",
        "remove_prompt": "Girdiyi kaldır?",
        "stop_disable": "Durdur+Devre Dışı Bırak",
        "skip_label": "Atla",
        "failed_save": "{ext} kaydedilemedi: {err}"
    },
    "de": {
        "choose_lang_header": "Sprache wählen / Select language:",
        "enter_number": "Geben Sie eine Zahl ein (1-10):",
        "invalid_choice": "Ungültige Auswahl. Versuchen Sie es erneut.",
        "installing_psutil": "psutil nicht gefunden — Installation wird automatisch durchgeführt...",
        "psutil_installed": "psutil installiert. Skript wird neu gestartet...",
        "psutil_failed": "Installation von psutil fehlgeschlagen: {err}",
        "scanning": "System wird auf mögliche Konflikte überprüft...",
        "results_short": "Ergebnisse (kurz):",
        "processes": "Prozesse",
        "startup": "Win32 StartupCommand",
        "hkcu": "HKCU Run",
        "services": "Dienste",
        "interactive_prompt": "Möchten Sie die interaktive Behebung starten?",
        "yes": "Ja",
        "no": "Nein — nur Bericht",
        "choose_report_formats": "Wählen Sie Berichtformate (durch Komma, Indizes):",
        "formats_list": "1:.txt 2:.json 3:.csv 4:.xml 5:.html 6:.md 7:.log 8:.yml 9:.ini 10:.pdf",
        "generating_reports": "Erzeuge Berichte in den ausgewählten Formaten...",
        "done": "Fertig.",
        "json_file": "JSON-Datei",
        "txt_file": "Textdatei",
        "press_enter": "Drücken Sie Enter zum Beenden...",
        "auto_install_reportlab": "Reportlab wird für PDF benötigt — Installation läuft...",
        "reportlab_failed": "Reportlab-Installation fehlgeschlagen: {err}",
        "pdf_created": "PDF erstellt: {path}",
        "action_prompt": "Aktion?",
        "kill": "Beenden",
        "skip": "Überspringen",
        "alternatives": "Alternativen",
        "check_hkcu": "HKCU prüfen",
        "remove": "Entfernen",
        "remove_prompt": "Eintrag entfernen?",
        "stop_disable": "Stoppen+Deaktivieren",
        "skip_label": "Überspringen",
        "failed_save": "Speichern von {ext} fehlgeschlagen: {err}"
    },
    "fr": {
        "choose_lang_header": "Choisir la langue / Select language:",
        "enter_number": "Entrez le numéro (1-10):",
        "invalid_choice": "Choix invalide. Réessayez.",
        "installing_psutil": "psutil introuvable — installation automatique...",
        "psutil_installed": "psutil installé. Redémarrage du script...",
        "psutil_failed": "Échec de l'installation de psutil: {err}",
        "scanning": "Analyse du système pour conflits potentiels...",
        "results_short": "Résultats (résumé):",
        "processes": "Processus",
        "startup": "Win32 StartupCommand",
        "hkcu": "HKCU Run",
        "services": "Services",
        "interactive_prompt": "Voulez-vous exécuter la correction interactive?",
        "yes": "Oui",
        "no": "Non — uniquement le rapport",
        "choose_report_formats": "Choisissez les formats de rapport (séparés par des virgules, indices):",
        "formats_list": "1:.txt 2:.json 3:.csv 4:.xml 5:.html 6:.md 7:.log 8:.yml 9:.ini 10:.pdf",
        "generating_reports": "Génération des rapports dans les formats sélectionnés...",
        "done": "Terminé.",
        "json_file": "Fichier JSON",
        "txt_file": "Fichier texte",
        "press_enter": "Appuyez sur Entrée pour quitter...",
        "auto_install_reportlab": "La génération PDF nécessite reportlab — installation...",
        "reportlab_failed": "Échec de l'installation de reportlab: {err}",
        "pdf_created": "PDF créé: {path}",
        "action_prompt": "Action?",
        "kill": "Terminer",
        "skip": "Ignorer",
        "alternatives": "Alternatives",
        "check_hkcu": "Vérifier HKCU",
        "remove": "Supprimer",
        "remove_prompt": "Supprimer l'entrée?",
        "stop_disable": "Arrêter+Désactiver",
        "skip_label": "Ignorer",
        "failed_save": "Échec de sauvegarde {ext}: {err}"
    },
    "it": {
        "choose_lang_header": "Scegli la lingua / Select language:",
        "enter_number": "Inserisci numero (1-10):",
        "invalid_choice": "Scelta non valida. Riprova.",
        "installing_psutil": "psutil non trovato — installazione automatica in corso...",
        "psutil_installed": "psutil installato. Riavvio dello script...",
        "psutil_failed": "Installazione di psutil fallita: {err}",
        "scanning": "Scansione del sistema per possibili conflitti...",
        "results_short": "Risultati (breve):",
        "processes": "Processi",
        "startup": "Win32 StartupCommand",
        "hkcu": "HKCU Run",
        "services": "Servizi",
        "interactive_prompt": "Vuoi eseguire la risoluzione interattiva?",
        "yes": "Sì",
        "no": "No — solo rapporto",
        "choose_report_formats": "Scegli i formati del rapporto (separati da virgola, indici):",
        "formats_list": "1:.txt 2:.json 3:.csv 4:.xml 5:.html 6:.md 7:.log 8:.yml 9:.ini 10:.pdf",
        "generating_reports": "Generazione rapporti nei formati scelti...",
        "done": "Fatto.",
        "json_file": "File JSON",
        "txt_file": "File di testo",
        "press_enter": "Premi Invio per uscire...",
        "auto_install_reportlab": "Per creare PDF è richiesto reportlab — installazione...",
        "reportlab_failed": "Installazione reportlab fallita: {err}",
        "pdf_created": "PDF creato: {path}",
        "action_prompt": "Azione?",
        "kill": "Chiudi",
        "skip": "Salta",
        "alternatives": "Alternative",
        "check_hkcu": "Controlla HKCU",
        "remove": "Rimuovi",
        "remove_prompt": "Rimuovere la voce?",
        "stop_disable": "Arresta+Disabilita",
        "skip_label": "Salta",
        "failed_save": "Salvataggio {ext} fallito: {err}"
    },
    "zh": {
        "choose_lang_header": "选择语言 / Select language:",
        "enter_number": "请输入数字 (1-10):",
        "invalid_choice": "选择无效。请重试。",
        "installing_psutil": "未找到 psutil — 正在自动安装...",
        "psutil_installed": "psutil 已安装。正在重启脚本...",
        "psutil_failed": "安装 psutil 失败: {err}",
        "scanning": "正在扫描系统以查找潜在冲突...",
        "results_short": "结果（简要）:",
        "processes": "进程",
        "startup": "Win32 启动命令",
        "hkcu": "HKCU Run",
        "services": "服务",
        "interactive_prompt": "是否运行交互式修复？",
        "yes": "是",
        "no": "否 — 仅报告",
        "choose_report_formats": "选择报告格式（用逗号分隔，编号）:",
        "formats_list": "1:.txt 2:.json 3:.csv 4:.xml 5:.html 6:.md 7:.log 8:.yml 9:.ini 10:.pdf",
        "generating_reports": "正在以所选格式生成报告...",
        "done": "完成。",
        "json_file": "JSON 文件",
        "txt_file": "文本文件",
        "press_enter": "按 Enter 退出...",
        "auto_install_reportlab": "生成 PDF 需要 reportlab — 正在安装...",
        "reportlab_failed": "安装 reportlab 失败: {err}",
        "pdf_created": "PDF 已创建: {path}",
        "action_prompt": "操作?",
        "kill": "终止",
        "skip": "跳过",
        "alternatives": "替代方案",
        "check_hkcu": "检查 HKCU",
        "remove": "删除",
        "remove_prompt": "删除条目？",
        "stop_disable": "停止并禁用",
        "skip_label": "跳过",
        "failed_save": "保存 {ext} 失败: {err}"
    },
    "ja": {
        "choose_lang_header": "言語を選択 / Select language:",
        "enter_number": "番号を入力してください (1-10):",
        "invalid_choice": "無効な選択です。もう一度お試しください。",
        "installing_psutil": "psutil が見つかりません — 自動インストール中...",
        "psutil_installed": "psutil がインストールされました。スクリプトを再起動します...",
        "psutil_failed": "psutil のインストールに失敗しました: {err}",
        "scanning": "潜在的な競合を検出するためにシステムをスキャンしています...",
        "results_short": "結果（簡易）:",
        "processes": "プロセス",
        "startup": "Win32 StartupCommand",
        "hkcu": "HKCU Run",
        "services": "サービス",
        "interactive_prompt": "対話型の修復を実行しますか？",
        "yes": "はい",
        "no": "いいえ — レポートのみ",
        "choose_report_formats": "レポート形式を選択（カンマ区切り、番号）：",
        "formats_list": "1:.txt 2:.json 3:.csv 4:.xml 5:.html 6:.md 7:.log 8:.yml 9:.ini 10:.pdf",
        "generating_reports": "選択した形式でレポートを生成しています...",
        "done": "完了。",
        "json_file": "JSON ファイル",
        "txt_file": "テキストファイル",
        "press_enter": "終了するには Enter キーを押してください...",
        "auto_install_reportlab": "PDF 作成には reportlab が必要です — インストール中...",
        "reportlab_failed": "reportlab のインストールに失敗しました: {err}",
        "pdf_created": "PDF 作成済み: {path}",
        "action_prompt": "操作?",
        "kill": "終了",
        "skip": "スキップ",
        "alternatives": "代替",
        "check_hkcu": "HKCU を確認",
        "remove": "削除",
        "remove_prompt": "エントリを削除しますか？",
        "stop_disable": "停止＋無効化",
        "skip_label": "スキップ",
        "failed_save": "{ext} の保存に失敗しました: {err}"
    }
}

# ---------------------------
# Helper: localized text
# ---------------------------
def choose_language() -> str:
    print("=" * 40)
    print("Language selection / Выбор языка / Selección de idioma")
    for i, (code, name) in enumerate(LANGS.items(), start=1):
        print(f"{i}. {name} ({code})")
    print("=" * 40)
    while True:
        choice = input("Enter number (1-10): ").strip()
        if not choice.isdigit():
            print("Please enter a number (1-10). / Введите номер (1-10).")
            continue
        idx = int(choice)
        if 1 <= idx <= len(LANGS):
            code = list(LANGS.keys())[idx - 1]
            return code
        print("Invalid choice. Try again.")

# Temporary set language selection
SELECTED_LANG = choose_language()
if SELECTED_LANG not in TRANSLATIONS:
    SELECTED_LANG = "en"

def t(key: str) -> str:
    return TRANSLATIONS.get(SELECTED_LANG, TRANSLATIONS["en"]).get(key, TRANSLATIONS["en"].get(key, key))

# ---------------------------
# Auto-install psutil (with restart)
# ---------------------------
def ensure_psutil():
    try:
        import psutil  # type: ignore
        return True
    except ImportError:
        print(t("installing_psutil"))
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "psutil"])
            print(t("psutil_installed"))
            # Restart with same args
            subprocess.Popen([sys.executable] + sys.argv)
            sys.exit(0)
        except Exception as e:
            print(t("psutil_failed").format(err=str(e)))
            return False

# ---------------------------
# PowerShell utility
# ---------------------------
def powershell_exec(cmd: str) -> Tuple[str, str, int]:
    try:
        p = subprocess.run(
            ['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', cmd],
            capture_output=True, text=True
        )
        return (p.stdout or "").strip(), (p.stderr or "").strip(), p.returncode
    except Exception as e:
        return "", str(e), 1

# ---------------------------
# Config & scanners (same logic as before)
# ---------------------------
SEARCH_TERMS = [
    "snip", "capture", "overlay", "hotkey", "shortcut", "screen",
    "obs", "discord", "steam", "onenote", "sharex", "autohotkey", "macro", "hook"
]

EDGE_LOCATIONS = [
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"
]

def get_edge_product_version() -> str:
    for path in EDGE_LOCATIONS:
        if os.path.exists(path):
            ps = f"(Get-Item '{path}').VersionInfo.ProductVersion"
            out, err, rc = powershell_exec(ps)
            if rc == 0 and out:
                return out
            return f"Error: {err or 'unknown'}"
    return "Not found"

def scan_running_processes() -> Any:
    try:
        import psutil  # type: ignore
    except Exception:
        return {"error": "psutil missing", "hint": "pip install psutil"}

    found = []
    for proc in psutil.process_iter(['pid', 'name', 'exe']):
        try:
            name = (proc.info.get('name') or "").lower()
            exe = (proc.info.get('exe') or "") or ""
            combined = f"{name} {exe}".lower()
            if any(term in combined for term in SEARCH_TERMS):
                found.append({
                    "name": proc.info.get('name'),
                    "pid": proc.info.get('pid'),
                    "path": exe
                })
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            continue
        except Exception as e:
            found.append({"error": f"Process scanning error: {e}"})
    return found

def scan_win32_startupcommand() -> Any:
    cmd = "Get-CimInstance -ClassName Win32_StartupCommand | Select-Object Name,Command | ConvertTo-Json -Depth 3"
    out, err, rc = powershell_exec(cmd)
    if rc != 0:
        return {"error": "WMI failed", "details": err}
    try:
        parsed = json.loads(out)
        items = parsed if isinstance(parsed, list) else [parsed]
        matches = []
        for it in items:
            name = (it.get('Name') or "").strip()
            command = (it.get('Command') or "").strip()
            if any(term in f"{name} {command}".lower() for term in SEARCH_TERMS):
                matches.append({"name": name, "command": command})
        return matches
    except json.JSONDecodeError:
        lines = out.splitlines()
        return [ln.strip() for ln in lines if any(term in ln.lower() for term in SEARCH_TERMS)]

def scan_hkcu_run_values() -> Any:
    cmd = r"Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run' | ConvertTo-Json -Depth 3"
    out, err, rc = powershell_exec(cmd)
    if rc != 0:
        return {"error": "HKCU read failed", "details": err}
    try:
        parsed = json.loads(out)
        results = []
        if isinstance(parsed, dict):
            for k, v in parsed.items():
                if k.lower().startswith("ps"):
                    continue
                value = str(v) if v is not None else ""
                if any(term in f"{k} {value}".lower() for term in SEARCH_TERMS):
                    results.append({"name": k, "value": value})
        elif isinstance(parsed, list):
            for item in parsed:
                for k, v in item.items():
                    if k.lower().startswith("ps"):
                        continue
                    value = str(v) if v is not None else ""
                    if any(term in f"{k} {value}".lower() for term in SEARCH_TERMS):
                        results.append({"name": k, "value": value})
        return results
    except json.JSONDecodeError:
        lines = out.splitlines()
        return [ln.strip() for ln in lines if any(term in ln.lower() for term in SEARCH_TERMS)]

def scan_windows_services() -> Any:
    cmd = "Get-CimInstance -ClassName Win32_Service | Select-Object Name,DisplayName,State,PathName | ConvertTo-Json -Depth 3"
    out, err, rc = powershell_exec(cmd)
    if rc != 0:
        return {"error": "Services fetch failed", "details": err}
    try:
        parsed = json.loads(out)
        items = parsed if isinstance(parsed, list) else [parsed]
        matches = []
        for svc in items:
            name = (svc.get('Name') or "").strip()
            display = (svc.get('DisplayName') or "").strip()
            state = (svc.get('State') or "").strip()
            path = (svc.get('PathName') or "").strip()
            if any(term in f"{name} {display} {path}".lower() for term in SEARCH_TERMS):
                matches.append({"name": name, "display_name": display, "state": state, "path": path})
        return matches
    except json.JSONDecodeError:
        lines = out.splitlines()
        return [ln.strip() for ln in lines if any(term in ln.lower() for term in SEARCH_TERMS)]

# ---------------------------
# Remediation actions
# ---------------------------
def kill_process_by_pid(pid: int) -> Dict[str, Any]:
    try:
        import psutil  # type: ignore
    except Exception:
        return {"ok": False, "error": "psutil missing"}
    try:
        proc = psutil.Process(pid)
        proc.terminate()
        proc.wait(timeout=5)
        return {"ok": True, "message": f"Process {pid} terminated"}
    except psutil.NoSuchProcess:
        return {"ok": False, "message": "Process no longer exists"}
    except psutil.AccessDenied:
        return {"ok": False, "message": "Access denied (admin required)"}
    except Exception as e:
        return {"ok": False, "message": str(e)}

def delete_hkcu_run_value(value_name: str) -> Dict[str, Any]:
    safe = value_name.replace("'", "''")
    cmd = f"Remove-ItemProperty -Path 'HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Run' -Name '{safe}' -ErrorAction Stop; 'OK'"
    out, err, rc = powershell_exec(cmd)
    if rc == 0 and "OK" in out:
        return {"ok": True, "message": f"Removed {value_name} from HKCU Run"}
    return {"ok": False, "message": err or out or "Unknown error"}

def stop_and_disable_service_by_name(svc_name: str) -> Dict[str, Any]:
    results = {"stop": None, "disable": None}
    try:
        p1 = subprocess.run(['sc', 'stop', svc_name], capture_output=True, text=True)
        results['stop'] = (p1.stdout or "") + (p1.stderr or "")
    except Exception as e:
        results['stop'] = str(e)
    try:
        p2 = subprocess.run(['sc', 'config', svc_name, 'start=', 'disabled'], capture_output=True, text=True)
        results['disable'] = (p2.stdout or "") + (p2.stderr or "")
    except Exception as e:
        results['disable'] = str(e)
    return results

# ---------------------------
# Localized prompts helper
# ---------------------------
def prompt_choice_localized(question: str, options: Dict[str, str]) -> str:
    keys_display = "/".join(options.keys())
    while True:
        print(question)
        for k, v in options.items():
            print(f"  [{k}] {v}")
        ans = input(f"{t('enter_number')} ({keys_display}): ").strip().lower()
        if ans in options:
            return ans
        print(t("invalid_choice"))

# ---------------------------
# Report save functions (10 formats)
# ---------------------------
def save_json(report: Dict[str, Any], path: str) -> str:
    with open(path, "w", encoding="utf-8") as f:
        json.dump(report, f, indent=4, ensure_ascii=False)
    return path

def save_txt(report: Dict[str, Any], path: str) -> str:
    lines = []
    lines.append("System Conflict Report")
    lines.append(f"Generated: {report.get('timestamp')}")
    lines.append("")
    sysinfo = report.get('system', {})
    lines.append(f"System: {sysinfo.get('os')} {sysinfo.get('release')} ({sysinfo.get('platform')})")
    lines.append(f"Edge: {report.get('edge_version')}")
    lines.append("")
    def dump_section(title: str, content: Any) -> None:
        lines.append("== " + title + " ==")
        if not content:
            lines.append("  (no entries)")
        elif isinstance(content, (list, tuple)):
            for item in content:
                lines.append("  " + json.dumps(item, ensure_ascii=False))
        else:
            lines.append("  " + json.dumps(content, ensure_ascii=False))
        lines.append("")
    dump_section("Process findings", report.get('process_conflicts'))
    dump_section("Win32 StartupCommand", report.get('startup_conflicts'))
    dump_section("HKCU Run values", report.get('hkcu_conflicts'))
    dump_section("Service findings", report.get('service_conflicts'))
    dump_section("Actions performed", report.get('actions'))
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    return path

def save_csv(report: Dict[str, Any], path: str) -> str:
    import csv
    rows = []
    rows.append(["section", "item_json"])
    for k in ("process_conflicts","startup_conflicts","hkcu_conflicts","service_conflicts","actions"):
        val = report.get(k)
        rows.append([k, json.dumps(val, ensure_ascii=False)])
    with open(path, "w", encoding="utf-8", newline='') as f:
        writer = csv.writer(f)
        writer.writerows(rows)
    return path

def save_xml(report: Dict[str, Any], path: str) -> str:
    from xml.etree.ElementTree import Element, SubElement, ElementTree
    root = Element("SystemConflictReport")
    meta = SubElement(root, "Generated")
    meta.text = str(report.get("timestamp"))
    system = SubElement(root, "System")
    for k,v in report.get("system", {}).items():
        el = SubElement(system, k)
        el.text = str(v)
    for section in ("process_conflicts","startup_conflicts","hkcu_conflicts","service_conflicts","actions"):
        sec_el = SubElement(root, section)
        items = report.get(section) or []
        if isinstance(items, dict):
            el = SubElement(sec_el, "item")
            el.text = json.dumps(items, ensure_ascii=False)
        else:
            for it in items:
                item_el = SubElement(sec_el, "item")
                item_el.text = json.dumps(it, ensure_ascii=False)
    ElementTree(root).write(path, encoding="utf-8", xml_declaration=True)
    return path

def save_html(report: Dict[str, Any], path: str) -> str:
    html = ["<html><head><meta charset='utf-8'><title>System Conflict Report</title></head><body>"]
    html.append(f"<h1>System Conflict Report</h1><p>Generated: {report.get('timestamp')}</p>")
    html.append("<pre>")
    html.append(json.dumps(report, indent=4, ensure_ascii=False))
    html.append("</pre></body></html>")
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(html))
    return path

def save_md(report: Dict[str, Any], path: str) -> str:
    lines = []
    lines.append("# System Conflict Report")
    lines.append(f"**Generated:** {report.get('timestamp')}\n")
    lines.append("```json")
    lines.append(json.dumps(report, indent=4, ensure_ascii=False))
    lines.append("```")
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    return path

def save_log(report: Dict[str, Any], path: str) -> str:
    return save_txt(report, path)

def save_yml(report: Dict[str, Any], path: str) -> str:
    try:
        import yaml  # type: ignore
    except Exception:
        with open(path, "w", encoding="utf-8") as f:
            f.write("# YAML-like dump\n")
            f.write(json.dumps(report, indent=2, ensure_ascii=False))
        return path
    with open(path, "w", encoding="utf-8") as f:
        yaml.safe_dump(report, f, allow_unicode=True)
    return path

def save_ini(report: Dict[str, Any], path: str) -> str:
    from configparser import ConfigParser
    cfg = ConfigParser()
    cfg["meta"] = {"generated": str(report.get("timestamp"))}
    sysinfo = report.get("system", {})
    cfg["system"] = {k: str(v) for k,v in sysinfo.items()}
    cfg["actions"] = {"data": json.dumps(report.get("actions") or [], ensure_ascii=False)}
    with open(path, "w", encoding="utf-8") as f:
        cfg.write(f)
    return path

def save_pdf(report: Dict[str, Any], path: str) -> str:
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.pdfgen import canvas
    except Exception:
        print(t("auto_install_reportlab"))
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "reportlab"])
        except Exception as e:
            print(t("reportlab_failed").format(err=str(e)))
            raise
        from reportlab.lib.pagesizes import A4
        from reportlab.pdfgen import canvas

    c = canvas.Canvas(path, pagesize=A4)
    width, height = A4
    text = c.beginText(40, height - 40)
    text.setFont("Helvetica", 10)
    text.textLine("System Conflict Report")
    text.textLine(f"Generated: {report.get('timestamp')}")
    text.textLine("")
    text.textLine("System info:")
    sysinfo = report.get("system", {})
    for k, v in sysinfo.items():
        text.textLine(f" {k}: {v}")
    text.textLine("")
    snippet = json.dumps(report, ensure_ascii=False, indent=2)
    for line in snippet.splitlines():
        if text.getY() < 60:
            c.drawText(text)
            c.showPage()
            text = c.beginText(40, height - 40)
            text.setFont("Helvetica", 10)
        text.textLine(line)
    c.drawText(text)
    c.save()
    return path

FORMAT_SAVE_FUNCS = {
    1: (".txt", save_txt),
    2: (".json", save_json),
    3: (".csv", save_csv),
    4: (".xml", save_xml),
    5: (".html", save_html),
    6: (".md", save_md),
    7: (".log", save_log),
    8: (".yml", save_yml),
    9: (".ini", save_ini),
    10: (".pdf", save_pdf)
}

# ---------------------------
# Main interactive flow
# ---------------------------
def main_flow():
    print(t("scanning"))
    detections = {
        "timestamp": datetime.now().isoformat(),
        "system": {
            "os": platform.system(),
            "release": platform.release(),
            "platform": platform.platform()
        },
        "edge_version": get_edge_product_version(),
        "process_conflicts": scan_running_processes(),
        "startup_conflicts": scan_win32_startupcommand(),
        "hkcu_conflicts": scan_hkcu_run_values(),
        "service_conflicts": scan_windows_services()
    }

    print(t("results_short"))
    def count_or_message(x: Any) -> str:
        if isinstance(x, dict) and x.get("error"):
            return f"Ошибка: {x.get('error')}"
        if isinstance(x, list):
            return str(len(x))
        return str(bool(x))

    print(f"  {t('processes')}: {count_or_message(detections['process_conflicts'])}")
    print(f"  {t('startup')}: {count_or_message(detections['startup_conflicts'])}")
    print(f"  {t('hkcu')}: {count_or_message(detections['hkcu_conflicts'])}")
    print(f"  {t('services')}: {count_or_message(detections['service_conflicts'])}")

    ans = prompt_choice_localized(t("interactive_prompt"), {"y": t("yes"), "n": t("no")})
    actions = []
    if ans == "y":
        pc = detections.get("process_conflicts") or []
        if isinstance(pc, dict) and pc.get("error"):
            print("Process scan:", pc)
        else:
            for proc in pc:
                name = proc.get("name"); pid = proc.get("pid"); path = proc.get("path")
                print(f"\n{name} (PID {pid})\n  {path}")
                ch = prompt_choice_localized(t("action_prompt"), {"k": t("kill"), "s": t("skip"), "a": t("alternatives")})
                if ch == "k":
                    res = kill_process_by_pid(pid)
                    actions.append({"action": "kill_process", "target": proc, "result": res})
                    print("Result:", res)
                else:
                    actions.append({"action": "skip_process", "target": proc})

        ac = detections.get("startup_conflicts") or []
        for e in ac:
            name = e.get("name"); cmd = e.get("command")
            print(f"\n{name}\n  {cmd}")
            ch = prompt_choice_localized(t("action_prompt"), {"i": t("check_hkcu"), "s": t("skip"), "a": t("alternatives")})
            if ch == "i":
                hk = detections.get("hkcu_conflicts") or []
                found = next((r for r in hk if r.get("name") == name), None)
                if found:
                    sub = prompt_choice_localized(t("remove_prompt"), {"y": t("yes"), "n": t("no")})
                    if sub == "y":
                        res = delete_hkcu_run_value(found.get("name"))
                        actions.append({"action": "delete_hkcu", "target": found, "result": res})
                else:
                    actions.append({"action": "manual_review", "target": e})
            else:
                actions.append({"action": "skip_autostart", "target": e})

        rc = detections.get("hkcu_conflicts") or []
        for r in rc:
            name = r.get("name"); val = r.get("value")
            print(f"\n{name}\n  {val}")
            ch = prompt_choice_localized(t("remove_prompt"), {"y": t("yes"), "n": t("no")})
            if ch == "y":
                res = delete_hkcu_run_value(name)
                actions.append({"action": "delete_hkcu", "target": r, "result": res})
            else:
                actions.append({"action": "skip_registry", "target": r})

        sc = detections.get("service_conflicts") or []
        for s in sc:
            name = s.get("name"); disp = s.get("display_name"); state = s.get("state")
            print(f"\n{name} ({disp}) state={state}")
            ch = prompt_choice_localized(t("action_prompt"), {"d": t("stop_disable"), "s": t("skip_label")})
            if ch == "d":
                res = stop_and_disable_service_by_name(name)
                actions.append({"action": "stop_disable_service", "target": s, "result": res})
            else:
                actions.append({"action": "skip_service", "target": s})
    else:
        print(t("done"))

    report = {
        "timestamp": detections["timestamp"],
        "system": detections["system"],
        "edge_version": detections["edge_version"],
        "process_conflicts": detections["process_conflicts"],
        "startup_conflicts": detections["startup_conflicts"],
        "hkcu_conflicts": detections["hkcu_conflicts"],
        "service_conflicts": detections["service_conflicts"],
        "actions": actions
    }

    print()
    print(t("choose_report_formats"))
    print(t("formats_list"))
    fmt_input = input(">>> ").strip()
    selected_indices: List[int] = []
    for token in fmt_input.split(","):
        token = token.strip()
        if not token:
            continue
        if token.isdigit():
            idx = int(token)
            if 1 <= idx <= 10:
                selected_indices.append(idx)
    if not selected_indices:
        selected_indices = [1, 2]

    print(t("generating_reports"))
    base = os.path.join(os.getcwd(), "system_conflict_report")
    for idx in selected_indices:
        ext, func = FORMAT_SAVE_FUNCS.get(idx, (None, None))
        if ext is None:
            continue
        path = base + ext
        try:
            func(report, path)
            if ext == ".json":
                print(f"{t('json_file')}: {path}")
            elif ext == ".txt":
                print(f"{t('txt_file')}: {path}")
            elif ext == ".pdf":
                print(t("pdf_created").format(path=path))
        except Exception as e:
            print(t("failed_save").format(ext=ext, err=str(e)))

    print(t("done"))
    try:
        input(t("press_enter"))
    except Exception:
        pass

# ---------------------------
# Run
# ---------------------------
if __name__ == "__main__":
    ok = ensure_psutil()
    # continue even if psutil installation failed (scans will show errors)
    main_flow()
