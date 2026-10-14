import json
import os
import sys

from sqlalchemy import create_engine
from sqlalchemy import text  
from sqlalchemy.engine import URL

def get_config_path():
    name = "db_config.json"
    if getattr(sys, "frozen", False):
      candidates = [os.path.join(os.path.dirname(sys.executable), name),
                    os.path.join(sys._MEIPASS, name)]
    else:
      candidates = [os.path.join(os.path.dirname(os.path.abspath(__file__)), name)]
    for p in candidates:
      if os.path.isfile(p):
        return p
    return candidates[0]

CONFIG_PATH = get_config_path() 
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
  _engines[name] = create_engine(url, pool_pre_ping=True, connect_args={"timeout": 5})
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
  sql += (f" ORDER BY [{FILTER_COLUMNS['branch']}],"
            f" [{FILTER_COLUMNS['department']}],"
            f" [{EMPLOYEE_FIELDS['Nama']}]")

  with get_engine(conn_name).connect() as conn:
    rows = conn.execute(text(sql), params).fetchall()
  return [["" if v is None else str(v).strip() for v in r] for r in rows]

def describe_db_error(e: Exception) -> str:
  if isinstance(e, FileNotFoundError):
    return str(e)
  if isinstance(e, KeyError):
    return f"Missing setting in db_config.json: {e}"

  low = str(getattr(e, "orig", None) or e).lower()
  if "login failed" in low:
    return "Wrong username or password for this database."
  if "data source name not found" in low or "can't open lib" in low:
    return "The SQL Server ODBC driver was not found. Check 'odbc_driver' in db_config.json."
  if "cannot open database" in low:
    return "The database name is wrong, or this login has no access to it."
  if any(s in low for s in ("timeout", "tcp provider", "network-related",
          "could not open a connection", "08001")):
    return "Cannot reach the server. Check the host and port, your network, or the VPN."
  return "Could not connect to the database."

def dispose_all():
  for e in _engines.values():
    e.dispose()
  _engines.clear()