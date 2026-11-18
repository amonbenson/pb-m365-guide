"""Builds teams/app-package.zip. Usage: python3 teams/build_package.py <public host name>"""
import json
import struct
import sys
import uuid
import zipfile
import zlib
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


def read_env():
    values = {}
    for line in (ROOT / ".env").read_text().splitlines():
        if "=" in line and not line.startswith("#"):
            key, value = line.split("=", 1)
            values[key.strip()] = value.strip()
    return values


def square_png(size, rgba, border=0):
    transparent = bytes((0, 0, 0, 0))
    def pixel(x, y):
        inside = border <= x < size - border and border <= y < size - border
        return transparent if border and inside else bytes(rgba)
    raw = b"".join(b"\x00" + b"".join(pixel(x, y) for x in range(size)) for y in range(size))
    def chunk(kind, data):
        return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data))
    header = struct.pack(">IIBBBBB", size, size, 8, 6, 0, 0, 0)
    return b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", header) + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b"")


def manifest(client_id, host):
    origin = f"https://{host}"
    return {
        "$schema": "https://developer.microsoft.com/json-schemas/teams/v1.22/MicrosoftTeams.schema.json",
        "manifestVersion": "1.22",
        "version": "1.0.0",
        "id": str(uuid.uuid5(uuid.NAMESPACE_URL, "teams-app:" + client_id)),
        "developer": {
            "name": "Entra Demo",
            "websiteUrl": origin,
            "privacyUrl": origin,
            "termsOfUseUrl": origin,
        },
        "name": {"short": "Entra Demo", "full": "PocketBase Entra ID demo"},
        "description": {
            "short": "PocketBase + Entra ID demo",
            "full": "Demo tab using Nested App Authentication with a PocketBase backend.",
        },
        "icons": {"color": "color.png", "outline": "outline.png"},
        "accentColor": "#4F6BED",
        "staticTabs": [
            {"entityId": "entra-demo-personal-tab", "name": "Entra Demo", "contentUrl": f"{origin}/index.html", "scopes": ["personal"]}
        ],
        "configurableTabs": [
            {"configurationUrl": f"{origin}/tab-config.html", "canUpdateConfiguration": True, "scopes": ["team", "groupChat"]}
        ],
        "permissions": ["identity"],
        "validDomains": [host],
        "webApplicationInfo": {"id": client_id, "resource": f"api://{client_id}"},
    }


def main():
    host = sys.argv[1]
    client_id = read_env()["ENTRA_CLIENT_ID"]
    output = ROOT / "teams" / "app-package.zip"
    with zipfile.ZipFile(output, "w") as archive:
        archive.writestr("manifest.json", json.dumps(manifest(client_id, host), indent=2))
        archive.writestr("color.png", square_png(192, (79, 107, 237, 255)))
        archive.writestr("outline.png", square_png(32, (255, 255, 255, 255), border=3))
    print(output)


main()
