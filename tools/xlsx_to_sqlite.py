#!/usr/bin/env python3
"""엑셀 DB(.xlsx) → 게시용 SQLite(.db) 변환 스크립트

사용법 (명령 프롬프트, 작업 폴더에서 실행):
    python tools\\xlsx_to_sqlite.py parking_db.xlsx data\\chatbot.db

- 먼저 convert_db.py의 검사를 모두 통과해야 합니다. 오류가 있으면 DB를 만들지 않습니다.
- 내부 메모(memo), 검토 필요 사항은 DB에 넣지 않습니다. (공개용 DB이기 때문)
- 게시(Y)된 FAQ와 주차장 정보만 들어갑니다.
"""
import json
import sqlite3
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import convert_db  # noqa: E402  (같은 폴더의 검사·변환 로직 재사용)

SCHEMA = """
CREATE TABLE meta (key TEXT PRIMARY KEY, value TEXT);
CREATE TABLE faq (
    faq_id TEXT PRIMARY KEY, center TEXT, category TEXT, question TEXT, answer TEXT,
    source TEXT, as_of TEXT, status TEXT
);
CREATE TABLE faq_place (
    faq_id TEXT, pos INTEGER, name TEXT, address TEXT, lot_type TEXT,
    PRIMARY KEY (faq_id, pos)
);
CREATE TABLE faq_synonym (
    keyword TEXT, faq_id TEXT,
    PRIMARY KEY (keyword, faq_id)
);
CREATE TABLE contact (pos INTEGER PRIMARY KEY, name TEXT, phone TEXT, hours TEXT, purpose TEXT);
CREATE TABLE center (pos INTEGER PRIMARY KEY, name TEXT, phone TEXT, hours TEXT, scope TEXT);
CREATE TABLE discount (pos INTEGER PRIMARY KEY, type TEXT, rate TEXT, proof TEXT);
CREATE TABLE menu (pos INTEGER PRIMARY KEY, menu_id TEXT UNIQUE, label TEXT, sub TEXT, icon TEXT, action TEXT);
CREATE TABLE menu_item (
    menu_id TEXT, pos INTEGER, faq_id TEXT,
    PRIMARY KEY (menu_id, pos)
);
"""


def build_json(src: Path) -> dict:
    """convert_db.py로 검사 후 임시 JSON을 만들고 읽어 온다. 오류면 None."""
    with tempfile.TemporaryDirectory() as tmp:
        tmp_json = Path(tmp) / "faq.json"
        code = convert_db.convert(src, tmp_json)
        if code != 0:
            return None
        return json.loads(tmp_json.read_text(encoding="utf-8"))


def main() -> int:
    if len(sys.argv) != 3:
        print("사용법: python tools\\xlsx_to_sqlite.py <엑셀.xlsx> <출력.db>")
        return 1
    src, out = Path(sys.argv[1]), Path(sys.argv[2])

    data = build_json(src)
    if data is None:
        print("[중단] 엑셀 검사에서 오류가 있어 DB를 만들지 않았습니다. 위 오류를 고친 뒤 다시 실행하세요.")
        return 1

    out.parent.mkdir(parents=True, exist_ok=True)
    if out.exists():
        out.unlink()
    con = sqlite3.connect(out)
    con.executescript(SCHEMA)

    con.execute("INSERT INTO meta VALUES ('generated_at', ?)", (data["generated_at"],))
    con.execute("INSERT INTO meta VALUES ('faq_count', ?)", (str(data["faq_count"]),))

    for f in data["faqs"]:
        con.execute("INSERT INTO faq VALUES (?,?,?,?,?,?,?,?)",
                    (f["id"], f.get("center", ""), f.get("category", ""), f["question"], f["answer"],
                     f["source"], f["as_of"], f["status"]))
        for pos, p in enumerate(f.get("places", []), start=1):
            con.execute("INSERT INTO faq_place VALUES (?,?,?,?,?)",
                        (f["id"], pos, p["name"], p["address"], p["lot_type"]))

    for keyword, ids in data["synonyms"].items():
        for fid in ids:
            con.execute("INSERT OR IGNORE INTO faq_synonym VALUES (?,?)", (keyword, fid))

    for pos, c in enumerate(data["contacts"], start=1):
        con.execute("INSERT INTO contact VALUES (?,?,?,?,?)",
                    (pos, c["name"], c["phone"], c["hours"], c["purpose"]))

    for pos, c in enumerate(data["centers"], start=1):
        con.execute("INSERT INTO center VALUES (?,?,?,?,?)",
                    (pos, c["name"], c["phone"], c["hours"], c["scope"]))

    for pos, d in enumerate(data["discounts"], start=1):
        con.execute("INSERT INTO discount VALUES (?,?,?,?)", (pos, d["type"], d["rate"], d["proof"]))

    for pos, m in enumerate(data["menus"], start=1):
        con.execute("INSERT INTO menu VALUES (?,?,?,?,?,?)",
                    (pos, m["id"], m["label"], m["sub"], m["icon"], m["action"]))
        for ipos, it in enumerate(m["items"], start=1):
            con.execute("INSERT INTO menu_item VALUES (?,?,?)", (m["id"], ipos, it["id"]))

    con.commit()
    con.close()
    print(f"DB 생성 완료: {out}")
    print(f"  FAQ {data['faq_count']}개, 메뉴 {len(data['menus'])}개, 할인 행 {len(data['discounts'])}개, "
          f"문의처 {len(data['contacts'])}개")
    return 0


if __name__ == "__main__":
    sys.exit(main())
