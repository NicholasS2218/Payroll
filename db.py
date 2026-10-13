import json
import os
import sys

from sqlalchemy import create_engine
from sqlalchemy import text  
from sqlalchemy.engine import URL

def get_config_dir():
  # next to the .exe when frozen, NOT inside sys._MEIPASS (that is a temp folder)
  if getattr(sys, "frozen", False):
      return os.path.dirname(sys.executable)
  return os.path.dirname(os.path.abspath(__file__))

CONFIG_PATH = os.path.join(get_config_dir(), "db_config.json")
_engines = {}

def read_config():
  if not os.path.isfile(CONFIG_PATH):
    raise FileNotFoundError(
      f"Database config not found:\n{CONFIG_PATH}\n\n"
      "Copy db_config.example.json to db_config.json and fill it in."
    )
  with open(CONFIG_PATH, encoding="utf-8") as f:
    return json.load(f)

def load_config():
  return read_config()["connections"]

def load_companies() -> dict:
  companies = read_config().get("companies", {})
  if not companies:
    raise ValueError("No 'companies' defined in db_config.json.")
  return companies

def get_engine(name: str = "MAIN"):
  if name in _engines:
    return _engines[name]
  
  connections = load_config()
  if name not in connections:
      raise ValueError(
          f"Unknown database connection '{name}'. "
          f"Defined in config: {', '.join(connections)}"
      )
  c = connections[name]

  url = URL.create(
    drivername=c["driver"],
    username=c["user"],
    password=c["password"],
    host=c["host"],
    port=c.get("port"),
    database=c["database"],
    query={
      "driver": c.get("odbc_driver", "ODBC Driver 17 for SQL Server"),
      "TrustServerCertificate": "yes",   # needed for Driver 18 with self-signed certs
    },
  )
  _engines[name] = create_engine(url, pool_pre_ping=True)
  return _engines[name]

EMPLOYEE_TABLE = "TB_Employee"
FILTER_COLUMNS = {
  "branch": "Branch",
  "department": "Department",
  "jabatan": "Jabatan",
  "group": "Grup",
}

def fetch_distinct(conn_name: str, key: str, filters: dict | None = None) -> list:
  col = FILTER_COLUMNS[key]
  where, params = [], {}
  for k, v in (filters or {}).items():
      if v and v != "All":
          where.append(f"[{FILTER_COLUMNS[k]}] = :{k}")
          params[k] = v

  sql = f"SELECT DISTINCT [{col}] FROM {EMPLOYEE_TABLE} WHERE [{col}] IS NOT NULL"
  if where:
      sql += " AND " + " AND ".join(where)
  sql += f" ORDER BY [{col}]"

  with get_engine(conn_name).connect() as conn:
      rows = conn.execute(text(sql), params).fetchall()
  return list(dict.fromkeys(str(r[0]).strip() for r in rows if str(r[0]).strip()))

EMPLOYEE_FIELDS = {    
  "Nama": "Name",
  "No NIK": "NIK",
  "Jabatan": "Jabatan",
  "Department": "Department",
}

def fetch_employees(conn_name: str, filters: dict) -> list:
  cols = ", ".join(f"[{c}]" for c in EMPLOYEE_FIELDS.values())
  where, params = [], {}
  for k, v in filters.items():
    if v and v != "All":
      where.append(f"[{FILTER_COLUMNS[k]}] = :{k}")
      params[k] = v

  sql = f"SELECT {cols} FROM {EMPLOYEE_TABLE}"
  if where:
    sql += " WHERE " + " AND ".join(where)
  sql += f" ORDER BY [{EMPLOYEE_FIELDS['Nama']}]"

  with get_engine(conn_name).connect() as conn:
    rows = conn.execute(text(sql), params).fetchall()
  return [["" if v is None else str(v).strip() for v in r] for r in rows]

def dispose_all():
  for e in _engines.values():
    e.dispose()
  _engines.clear()