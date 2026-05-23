"""
Isolated ScreenPlant scale check.

Creates a temporary SQLite database with the same hot-path submission indexes
used by the app, inserts synthetic submissions, and times the common operations
that matter for 30k-40k student deployments.

It does not import app.py and does not touch production/local ScreenPlant data.
"""

import argparse
import hashlib
import json
import os
import sqlite3
import tempfile
import time
from datetime import datetime, timedelta


def identity_hash(school_id, name, roll_no, class_dept):
    basis = "|".join([school_id, name.lower(), roll_no.lower(), class_dept.lower()])
    return hashlib.sha256(basis.encode("utf-8")).hexdigest()


def create_schema(conn):
    conn.execute("""
        CREATE TABLE screenplant_schools (
            id TEXT PRIMARY KEY,
            name TEXT,
            school_type TEXT,
            created_at TEXT,
            updated_at TEXT NOT NULL,
            data TEXT NOT NULL
        )
    """)
    conn.execute("""
        CREATE TABLE screenplant_submissions (
            id TEXT PRIMARY KEY,
            school_id TEXT,
            unique_id TEXT,
            name TEXT,
            roll_no TEXT,
            class_dept TEXT,
            status TEXT,
            submitted_at TEXT,
            updated_at TEXT NOT NULL,
            data TEXT NOT NULL
        )
    """)
    conn.execute("CREATE INDEX idx_screenplant_submissions_school ON screenplant_submissions (school_id)")
    conn.execute("CREATE INDEX idx_screenplant_submissions_status ON screenplant_submissions (status)")
    conn.execute("CREATE INDEX idx_screenplant_submissions_identity ON screenplant_submissions (school_id, json_extract(data, '$.submission_identity_hash'))")
    conn.commit()


def seed(conn, students, schools):
    now = datetime.now()
    school_ids = [f"sch_{i:03d}" for i in range(schools)]
    conn.executemany(
        "INSERT INTO screenplant_schools (id, name, school_type, created_at, updated_at, data) VALUES (?, ?, ?, ?, ?, ?)",
        [
            (
                sid,
                f"School {idx + 1}",
                "School",
                now.isoformat(),
                now.isoformat(),
                json.dumps({"id": sid, "name": f"School {idx + 1}", "type": "School"}),
            )
            for idx, sid in enumerate(school_ids)
        ],
    )

    rows = []
    for i in range(students):
        sid = school_ids[i % schools]
        roll = f"{sid.upper()}-{i:05d}"
        name = f"Student {i:05d}"
        cls = f"{1 + (i % 12)} - {'ABCD'[i % 4]}"
        submitted_at = (now - timedelta(seconds=i)).isoformat()
        sub = {
            "id": f"sub_{i:06d}",
            "school_id": sid,
            "name": name,
            "roll_no": roll,
            "unique_id": roll,
            "class_dept": cls,
            "status": "pending" if i % 5 else "generated",
            "photo_path": f"/r2/photos/{roll}.jpg",
            "submitted_at": submitted_at,
            "submission_identity_hash": identity_hash(sid, name, roll, cls),
        }
        rows.append(
            (
                sub["id"],
                sid,
                roll,
                name,
                roll,
                cls,
                sub["status"],
                submitted_at,
                now.isoformat(),
                json.dumps(sub),
            )
        )
    conn.executemany(
        """
        INSERT INTO screenplant_submissions
            (id, school_id, unique_id, name, roll_no, class_dept, status, submitted_at, updated_at, data)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        rows,
    )
    conn.commit()
    return school_ids


def timed(label, fn):
    start = time.perf_counter()
    result = fn()
    elapsed = (time.perf_counter() - start) * 1000
    print(f"{label}: {elapsed:.1f} ms")
    return result


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--students", type=int, default=40000)
    parser.add_argument("--schools", type=int, default=40)
    args = parser.parse_args()

    fd, path = tempfile.mkstemp(prefix="screenplant_scale_", suffix=".sqlite3")
    os.close(fd)
    try:
        conn = sqlite3.connect(path)
        create_schema(conn)
        school_ids = timed("seed synthetic rows", lambda: seed(conn, args.students, args.schools))
        sample_school = school_ids[len(school_ids) // 2]

        timed(
            "count one school",
            lambda: conn.execute(
                "SELECT COUNT(*) FROM screenplant_submissions WHERE school_id = ?",
                (sample_school,),
            ).fetchone()[0],
        )
        timed(
            "recent page one school",
            lambda: conn.execute(
                """
                SELECT data FROM screenplant_submissions
                WHERE school_id = ?
                ORDER BY submitted_at DESC, id DESC
                LIMIT 50 OFFSET 0
                """,
                (sample_school,),
            ).fetchall(),
        )
        duplicate_hash = identity_hash(sample_school, "Student 00020", f"{sample_school.upper()}-00020", "9 - A")
        timed(
            "duplicate identity check",
            lambda: conn.execute(
                """
                SELECT COUNT(*) FROM screenplant_submissions
                WHERE school_id = ?
                AND json_extract(data, '$.submission_identity_hash') = ?
                """,
                (sample_school, duplicate_hash),
            ).fetchone()[0],
        )
        timed(
            "insert one new submission",
            lambda: conn.execute(
                """
                INSERT INTO screenplant_submissions
                    (id, school_id, unique_id, name, roll_no, class_dept, status, submitted_at, updated_at, data)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    "sub_new",
                    sample_school,
                    "NEW-001",
                    "New Student",
                    "NEW-001",
                    "10 - A",
                    "pending",
                    datetime.now().isoformat(),
                    datetime.now().isoformat(),
                    json.dumps({"id": "sub_new", "school_id": sample_school, "submission_identity_hash": "new"}),
                ),
            ),
        )
        conn.commit()
        print(f"temporary database: {path}")
    finally:
        try:
            conn.close()
        except Exception:
            pass
        try:
            os.remove(path)
        except OSError:
            pass


if __name__ == "__main__":
    main()
