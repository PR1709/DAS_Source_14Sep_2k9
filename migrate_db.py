#!/usr/bin/env python3
import os
import struct
import re
from datetime import datetime, date

def clean_name(name):
    """Clean table or column names to be PostgreSQL compatible."""
    name = name.strip().lower()
    # Replace non-alphanumeric characters with underscores
    name = re.sub(r'[^a-z0-9_]', '_', name)
    # Ensure it doesn't start with a number
    if name and name[0].isdigit():
        name = "_" + name
    # Avoid PostgreSQL reserved keywords (simplified list)
    reserved = {"user", "table", "limit", "order", "select", "insert", "update", "delete", "from", "where", "group", "by"}
    if name in reserved:
        name = f"_{name}"
    return name

class DBFReader:
    """A simple, dependency-free reader for dBase / Visual FoxPro .dbf and .fpt files."""
    def __init__(self, dbf_path):
        self.dbf_path = dbf_path
        self.fpt_path = os.path.splitext(dbf_path)[0] + ".fpt"
        self.has_memo = os.path.exists(self.fpt_path)
        self.fields = []
        self.records = []
        self.memo_block_size = 512
        
        if self.has_memo:
            self._read_memo_header()
        self._read_dbf()

    def _read_memo_header(self):
        try:
            with open(self.fpt_path, "rb") as f:
                # First 4 bytes: next free block
                # Next 2 bytes: unused
                # Bytes 6-7: block size (big-endian)
                f.seek(6)
                block_size = struct.unpack(">H", f.read(2))[0]
                if block_size > 0:
                    self.memo_block_size = block_size
        except Exception as e:
            print(f"Warning: Could not read memo header for {self.fpt_path}: {e}")

    def _read_memo_val(self, block_num):
        if not block_num or block_num == 0:
            return ""
        try:
            with open(self.fpt_path, "rb") as f:
                f.seek(block_num * self.memo_block_size)
                # VFP Memo block structure:
                # Bytes 0-3: signature (1 = text, 0 = picture/binary)
                # Bytes 4-7: length of memo (big endian)
                sig_bytes = f.read(4)
                if not sig_bytes:
                    return ""
                sig = struct.unpack(">I", sig_bytes)[0]
                length = struct.unpack(">I", f.read(4))[0]
                data = f.read(length)
                if sig == 1:
                    return data.decode("latin1", errors="replace").strip()
                else:
                    return f"\\x{data.hex()}" # Return hex representation for binary/picture
        except Exception as e:
            return f"[Error reading memo block {block_num}: {e}]"

    def _read_dbf(self):
        with open(self.dbf_path, "rb") as f:
            # Header
            header = f.read(32)
            if len(header) < 32:
                return
            
            num_records, header_len, record_len = struct.unpack("<IHH", header[4:12])
            
            # Read field descriptors
            # Each descriptor is 32 bytes. Loop until we hit 0x0D (13)
            while True:
                next_byte = f.read(1)
                if not next_byte or next_byte[0] == 0x0D:
                    break
                descriptor = next_byte + f.read(31)
                if len(descriptor) < 32:
                    break
                
                name = descriptor[:11].split(b'\x00')[0].decode("latin1", errors="replace")
                field_type = chr(descriptor[11])
                length = descriptor[16]
                decimals = descriptor[17]
                
                self.fields.append({
                    "name": name,
                    "type": field_type,
                    "length": length,
                    "decimals": decimals
                })
            
            # Move to the start of records
            f.seek(header_len)
            
            # Read records
            for _ in range(num_records):
                record_data = f.read(record_len)
                if len(record_data) < record_len:
                    break
                
                # First byte is deletion flag (space ' ' means active, '*' means deleted)
                is_deleted = record_data[0] == 0x2A # '*'
                if is_deleted:
                    continue
                
                record = {}
                offset = 1
                for field in self.fields:
                    val_bytes = record_data[offset:offset+field["length"]]
                    offset += field["length"]
                    
                    ftype = field["type"]
                    val = ""
                    
                    if ftype == 'C': # Character
                        val = val_bytes.decode("latin1", errors="replace").strip()
                    elif ftype in ('N', 'F'): # Numeric / Float
                        val_str = val_bytes.decode("latin1", errors="replace").strip()
                        if not val_str:
                            val = None
                        elif '.' in val_str:
                            val = float(val_str)
                        else:
                            val = int(val_str) if val_str else 0
                    elif ftype == 'L': # Logical
                        c = chr(val_bytes[0]) if val_bytes else '?'
                        if c in ('Y', 'y', 'T', 't', '1'):
                            val = True
                        elif c in ('N', 'n', 'F', 'f', '0'):
                            val = False
                        else:
                            val = None
                    elif ftype == 'D': # Date (YYYYMMDD)
                        val_str = val_bytes.decode("latin1", errors="replace").strip()
                        if val_str and val_str != "00000000":
                            try:
                                val = datetime.strptime(val_str, "%Y%m%d").date()
                            except:
                                val = val_str
                        else:
                            val = None
                    elif ftype == 'M': # Memo
                        if len(val_bytes) == 4:
                            block_num = struct.unpack("<I", val_bytes)[0]
                        else:
                            val_str = val_bytes.decode("latin1", errors="replace").strip()
                            block_num = int(val_str) if val_str.isdigit() else 0
                        val = self._read_memo_val(block_num)
                    elif ftype == 'I': # Integer
                        val = struct.unpack("<i", val_bytes)[0]
                    elif ftype == 'T': # DateTime
                        # 8 bytes: two 32-bit integers (Julian day, milliseconds since midnight)
                        if len(val_bytes) == 8:
                            julian, msecs = struct.unpack("<ii", val_bytes)
                            # Julian day 0 is Nov 24, 4714 BC. VFP uses standard Julian days.
                            # We can convert Julian day to a timestamp
                            if julian > 0:
                                # Standard formula or simple offset
                                # Julian day 2440588 is Jan 1, 1970
                                epoch_days = julian - 2440588
                                ts = epoch_days * 86400 + (msecs / 1000.0)
                                try:
                                    val = datetime.fromtimestamp(ts)
                                except:
                                    val = f"Julian Day: {julian}, Time: {msecs}ms"
                            else:
                                val = None
                        else:
                            val = None
                    else:
                        val = val_bytes.decode("latin1", errors="replace").strip()
                    
                    record[field["name"]] = val
                
                self.records.append(record)

def map_type_to_postgres(field):
    ftype = field["type"]
    length = field["length"]
    decimals = field["decimals"]
    
    if ftype == 'C':
        return f"VARCHAR({length})"
    elif ftype == 'N':
        if decimals > 0:
            return f"NUMERIC({length}, {decimals})"
        elif length > 9:
            return "BIGINT"
        else:
            return "INTEGER"
    elif ftype == 'F':
        return "DOUBLE PRECISION"
    elif ftype == 'I':
        return "INTEGER"
    elif ftype == 'L':
        return "BOOLEAN"
    elif ftype == 'D':
        return "DATE"
    elif ftype == 'T':
        return "TIMESTAMP"
    elif ftype == 'M':
        return "TEXT"
    elif ftype == 'G':
        return "BYTEA"
    else:
        return "TEXT"

def parse_source_files(origin_dir):
    """Scan source files (.frm, .bas, .cls, .prg) for SQL queries, table references, and connection details."""
    extracted_data = []
    
    # Common SQL query regex patterns
    sql_patterns = [
        re.compile(r'(SELECT\s+.*?\s+FROM\s+\w+)', re.IGNORECASE | re.DOTALL),
        re.compile(r'(INSERT\s+INTO\s+\w+.*?\))', re.IGNORECASE | re.DOTALL),
        re.compile(r'(UPDATE\s+\w+\s+SET\s+.*?)', re.IGNORECASE | re.DOTALL),
        re.compile(r'(DELETE\s+FROM\s+\w+)', re.IGNORECASE | re.DOTALL),
    ]
    
    connection_patterns = [
        re.compile(r'(Provider=.*?;)', re.IGNORECASE),
        re.compile(r'(Driver=.*?;)', re.IGNORECASE),
        re.compile(r'(ConnectionString\s*=.*)', re.IGNORECASE),
    ]

    for root, _, files in os.walk(origin_dir):
        for file in files:
            ext = os.path.splitext(file)[1].lower()
            if ext in ('.frm', '.bas', '.cls', '.prg'):
                filepath = os.path.join(root, file)
                try:
                    with open(filepath, 'r', encoding='latin1') as f:
                        content = f.read()
                    
                    file_queries = []
                    file_conns = []
                    
                    # Search SQL
                    for pattern in sql_patterns:
                        for match in pattern.finditer(content):
                            # Clean up whitespace and limit length
                            query = re.sub(r'\s+', ' ', match.group(1)).strip()
                            if len(query) > 10 and query not in file_queries:
                                file_queries.append(query[:200])
                    
                    # Search Connection strings
                    for pattern in connection_patterns:
                        for match in pattern.finditer(content):
                            conn = match.group(1).strip()
                            if conn not in file_conns:
                                file_conns.append(conn)
                                
                    if file_queries or file_conns:
                        extracted_data.append({
                            "file": os.path.basename(filepath),
                            "queries": file_queries,
                            "connections": file_conns
                        })
                except Exception as e:
                    print(f"Error parsing file {file}: {e}")
                    
    return extracted_data

def generate_postgres_dump(origin_dir, output_sql_path):
    sql_lines = []
    sql_lines.append("-- ===========================================================================")
    sql_lines.append("-- PostgreSQL Migration Script")
    sql_lines.append(f"-- Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    sql_lines.append(f"-- Source Directory: {origin_dir}")
    sql_lines.append("-- ===========================================================================\n")
    
    # 1. Parse source files for documentation / metadata
    print("Scanning source files (.frm, .bas, .cls, .prg) for SQL and connection metadata...")
    metadata = parse_source_files(origin_dir)
    if metadata:
        sql_lines.append("-- ===========================================================================")
        sql_lines.append("-- EXTRACTED LEGACY APPLICATION METADATA & SQL QUERIES")
        sql_lines.append("-- ===========================================================================")
        for item in metadata:
            sql_lines.append(f"-- File: {item['file']}")
            if item['connections']:
                sql_lines.append("--   Connection Strings / Drivers:")
                for conn in item['connections']:
                    sql_lines.append(f"--     {conn}")
            if item['queries']:
                sql_lines.append("--   Detected SQL Queries:")
                for q in item['queries']:
                    sql_lines.append(f"--     {q}")
            sql_lines.append("-- ---------------------------------------------------------------------------")
        sql_lines.append("\n")

    # 2. Process all DBF files
    print("Processing database files (.dbf)...")
    dbf_files = [f for f in os.listdir(origin_dir) if f.lower().endswith('.dbf')]
    dbf_files.sort()
    
    for dbf_file in dbf_files:
        dbf_path = os.path.join(origin_dir, dbf_file)
        table_name = clean_name(os.path.splitext(dbf_file)[0])
        print(f"  Reading {dbf_file} -> Table: {table_name}")
        
        try:
            reader = DBFReader(dbf_path)
            if not reader.fields:
                print(f"    Warning: No fields found in {dbf_file}")
                continue
                
            sql_lines.append(f"-- Table: {table_name} (Migrated from {dbf_file})")
            sql_lines.append(f"DROP TABLE IF EXISTS {table_name} CASCADE;")
            sql_lines.append(f"CREATE TABLE {table_name} (")
            
            field_lines = []
            for field in reader.fields:
                col_name = clean_name(field["name"])
                col_type = map_type_to_postgres(field)
                field_lines.append(f"    {col_name} {col_type}")
            
            sql_lines.append(",\n".join(field_lines))
            sql_lines.append(");\n")
            
            # Generate INSERTS
            if reader.records:
                sql_lines.append(f"-- Inserting {len(reader.records)} records into {table_name}")
                for record in reader.records:
                    cols = []
                    vals = []
                    for field in reader.fields:
                        col_name = clean_name(field["name"])
                        val = record[field["name"]]
                        
                        cols.append(col_name)
                        if val is None:
                            vals.append("NULL")
                        elif isinstance(val, bool):
                            vals.append("TRUE" if val else "FALSE")
                        elif isinstance(val, (int, float)):
                            vals.append(str(val))
                        elif isinstance(val, datetime):
                            vals.append(f"'{val.strftime('%Y-%m-%d %H:%M:%S')}'")
                        elif isinstance(val, date):
                            vals.append(f"'{val.strftime('%Y-%m-%d')}'")
                        else:
                            # Escape single quotes
                            escaped_val = str(val).replace("'", "''")
                            vals.append(f"'{escaped_val}'")
                            
                    cols_str = ", ".join(cols)
                    vals_str = ", ".join(vals)
                    sql_lines.append(f"INSERT INTO {table_name} ({cols_str}) VALUES ({vals_str});")
                sql_lines.append("\n")
            else:
                sql_lines.append(f"-- No records found in {dbf_file}\n")
                
        except Exception as e:
            print(f"    Error processing {dbf_file}: {e}")
            sql_lines.append(f"-- Error migrating {dbf_file}: {e}\n")

    # Write output SQL file
    with open(output_sql_path, "w", encoding="utf-8") as f:
        f.write("\n".join(sql_lines))
    print(f"\nSuccessfully generated PostgreSQL initialization dump at: {output_sql_path}")

if __name__ == "__main__":
    base_dir = os.path.dirname(os.path.abspath(__file__))
    origin_dir = os.path.join(base_dir, "Origin")
    dest_dir = os.path.join(base_dir, "MigrateDest_v1")
    os.makedirs(dest_dir, exist_ok=True)
    output_sql = os.path.join(dest_dir, "init_postgres.sql")
    
    if not os.path.exists(origin_dir):
        print(f"Error: Origin directory not found at {origin_dir}")
    else:
        generate_postgres_dump(origin_dir, output_sql)
