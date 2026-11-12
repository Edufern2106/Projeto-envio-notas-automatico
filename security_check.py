"""
Verificação de segurança local — DS Unificador OCR
Execute: python security_check.py
"""

import os
import re
import sys
from collections import Counter, defaultdict
from datetime import datetime, timedelta
from pathlib import Path

LOG_FILE = Path(__file__).parent / "backend" / "access.log"

# IPs considerados locais/seguros
LOCAL_IPS = {"127.0.0.1", "::1", "localhost"}

# Limites de alerta
MAX_REQUESTS_PER_HOUR_PER_IP = 200
MAX_ERRORS_PER_HOUR = 50


def parse_log():
    if not LOG_FILE.exists():
        return []
    entries = []
    pat = re.compile(
        r"(\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}) \| (\S+) \| (\w+) (.+?) \| (\d{3}) \| (\d+)ms \| (.+)"
    )
    with open(LOG_FILE, encoding="utf-8") as f:
        for line in f:
            m = pat.match(line.strip())
            if m:
                entries.append({
                    "ts":     datetime.fromisoformat(m.group(1)),
                    "ip":     m.group(2),
                    "method": m.group(3),
                    "path":   m.group(4),
                    "status": int(m.group(5)),
                    "ms":     int(m.group(6)),
                    "ua":     m.group(7),
                })
    return entries


def check_config():
    print("=" * 60)
    print("1. VERIFICAÇÃO DE CONFIGURAÇÕES DE SEGURANÇA")
    print("=" * 60)
    main_py = Path(__file__).parent / "backend" / "main.py"
    if not main_py.exists():
        print("  ✗ backend/main.py não encontrado!")
        return False

    src = main_py.read_text(encoding="utf-8")
    checks = {
        "CORS não usa wildcard (*)":        'allow_origins=["*"]' not in src,
        "Validação de magic bytes":          "_PDF_MAGIC" in src and "_ZIP_MAGIC" in src,
        "Limite de arquivos por envio":      "MAX_FILES_PER_REQUEST" in src,
        "Proteção contra ZIP bomb":          "MAX_ZIP_MEMBERS" in src,
        "Limpeza automática de jobs":        "JOB_TTL_SECONDS" in src,
        "Swagger desabilitado":              "docs_url=None" in src,
        "Log de acesso ativo":               "access_log_middleware" in src,
        "Erros sem stack trace":             "Erro interno no processamento" in src,
    }
    all_ok = True
    for name, ok in checks.items():
        icon = "✓" if ok else "✗ ALERTA"
        print(f"  {icon}  {name}")
        if not ok:
            all_ok = False
    return all_ok


def check_access_log():
    print("\n" + "=" * 60)
    print("2. ANÁLISE DO LOG DE ACESSOS")
    print("=" * 60)

    entries = parse_log()
    if not entries:
        print("  Nenhuma entrada no log ainda.")
        return True

    now = datetime.now()
    last_24h = [e for e in entries if now - e["ts"] < timedelta(hours=24)]
    last_1h  = [e for e in entries if now - e["ts"] < timedelta(hours=1)]

    print(f"  Total de requisições no log:    {len(entries)}")
    print(f"  Últimas 24 horas:               {len(last_24h)}")
    print(f"  Última hora:                    {len(last_1h)}")

    alerts = []

    # IPs externos
    ext_ips = {e["ip"] for e in last_24h if e["ip"] not in LOCAL_IPS}
    if ext_ips:
        alerts.append(f"IPs EXTERNOS acessaram nas últimas 24h: {', '.join(sorted(ext_ips))}")

    # Volume por IP na última hora
    ip_count = Counter(e["ip"] for e in last_1h)
    for ip, count in ip_count.items():
        if count > MAX_REQUESTS_PER_HOUR_PER_IP:
            alerts.append(f"Volume suspeito: {ip} fez {count} requisições na última hora")

    # Erros na última hora
    errors = [e for e in last_1h if e["status"] >= 400]
    if len(errors) > MAX_ERRORS_PER_HOUR:
        alerts.append(f"Alto volume de erros na última hora: {len(errors)} respostas 4xx/5xx")

    # Tentativas em endpoints inexistentes
    not_found = [e for e in last_24h if e["status"] == 404 and not e["path"].startswith("/api/")]
    if len(not_found) > 10:
        paths = Counter(e["path"] for e in not_found).most_common(5)
        alerts.append(f"Possível varredura: {len(not_found)} requisições 404 fora da API: {paths}")

    # Upload rejeitado (magic bytes / tipo inválido)
    rejections = [e for e in last_24h if e["status"] == 400 and e["path"] == "/api/process"]
    if rejections:
        print(f"\n  Uploads rejeitados (últimas 24h): {len(rejections)}")

    # Resumo de endpoints acessados
    print("\n  Endpoints mais acessados (últimas 24h):")
    for path, count in Counter(e["path"] for e in last_24h).most_common(5):
        print(f"    {count:4d}×  {path}")

    # Alertas
    print()
    if alerts:
        print("  ⚠️  ALERTAS ENCONTRADOS:")
        for a in alerts:
            print(f"     • {a}")
        return False
    else:
        print("  ✓  Nenhum acesso suspeito detectado.")
        return True


def check_suspicious_files():
    print("\n" + "=" * 60)
    print("3. ARQUIVOS SUSPEITOS NO REPOSITÓRIO")
    print("=" * 60)
    root = Path(__file__).parent
    # .bat/.cmd/.ps1 são normais para scripts de startup Windows
    suspicious_exts = {".exe", ".dll", ".vbs", ".msi", ".scr", ".pif"}
    found = []
    for p in root.rglob("*"):
        if any(str(p).startswith(str(root / skip)) for skip in
               [".git", "frontend/node_modules", "frontend/dist"]):
            continue
        if p.suffix.lower() in suspicious_exts:
            found.append(str(p.relative_to(root)))

    if found:
        print("  ⚠️  ARQUIVOS SUSPEITOS ENCONTRADOS:")
        for f in found:
            print(f"     • {f}")
        return False
    else:
        print("  ✓  Nenhum arquivo suspeito encontrado.")
        return True


def main():
    print(f"\nDS Unificador OCR — Verificação de Segurança")
    print(f"Data/Hora: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")

    r1 = check_config()
    r2 = check_access_log()
    r3 = check_suspicious_files()

    print("\n" + "=" * 60)
    if r1 and r2 and r3:
        print("RESULTADO: ✓ SISTEMA SEGURO — Nenhum problema encontrado.")
        sys.exit(0)
    else:
        print("RESULTADO: ⚠️  ATENÇÃO — Verifique os alertas acima.")
        sys.exit(1)


if __name__ == "__main__":
    main()
