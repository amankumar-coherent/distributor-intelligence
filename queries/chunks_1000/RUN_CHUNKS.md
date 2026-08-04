# Run 1000 markets as 10 chunks (from D:\)

Project root: `D:\vendor-intel`  
Chunk files: `D:\vendor-intel\queries\chunks_1000\` (also copied to `D:\chunks_1000\`)

## Before you start

1. Stop any old mega-run on C: (Ctrl+C).
2. Keep SearXNG up. Start LinkedIn MCP once if you need Founded/HQ.
3. On **one PC**, prefer **2–3 terminals at a time**, not all 10.

## Each terminal — first lines

```powershell
cd D:\vendor-intel
$env:PYTHONPATH = ".\src"
$env:PYTHONIOENCODING = "utf-8"
```

## Terminal 1–10 commands

**Terminal 1**
```powershell
.\.venv\Scripts\python.exe run_batch.py --input queries\chunks_1000\market_chunk_01.txt --batch-id markets-1000-chunk-01 --mode quality --max-concurrent 3 --no-apollo --no-name-resolve --yes
```

**Terminal 2**
```powershell
.\.venv\Scripts\python.exe run_batch.py --input queries\chunks_1000\market_chunk_02.txt --batch-id markets-1000-chunk-02 --mode quality --max-concurrent 3 --no-apollo --no-name-resolve --yes
```

**Terminal 3**
```powershell
.\.venv\Scripts\python.exe run_batch.py --input queries\chunks_1000\market_chunk_03.txt --batch-id markets-1000-chunk-03 --mode quality --max-concurrent 3 --no-apollo --no-name-resolve --yes
```

**Terminal 4**
```powershell
.\.venv\Scripts\python.exe run_batch.py --input queries\chunks_1000\market_chunk_04.txt --batch-id markets-1000-chunk-04 --mode quality --max-concurrent 3 --no-apollo --no-name-resolve --yes
```

**Terminal 5**
```powershell
.\.venv\Scripts\python.exe run_batch.py --input queries\chunks_1000\market_chunk_05.txt --batch-id markets-1000-chunk-05 --mode quality --max-concurrent 3 --no-apollo --no-name-resolve --yes
```

**Terminal 6**
```powershell
.\.venv\Scripts\python.exe run_batch.py --input queries\chunks_1000\market_chunk_06.txt --batch-id markets-1000-chunk-06 --mode quality --max-concurrent 3 --no-apollo --no-name-resolve --yes
```

**Terminal 7**
```powershell
.\.venv\Scripts\python.exe run_batch.py --input queries\chunks_1000\market_chunk_07.txt --batch-id markets-1000-chunk-07 --mode quality --max-concurrent 3 --no-apollo --no-name-resolve --yes
```

**Terminal 8**
```powershell
.\.venv\Scripts\python.exe run_batch.py --input queries\chunks_1000\market_chunk_08.txt --batch-id markets-1000-chunk-08 --mode quality --max-concurrent 3 --no-apollo --no-name-resolve --yes
```

**Terminal 9**
```powershell
.\.venv\Scripts\python.exe run_batch.py --input queries\chunks_1000\market_chunk_09.txt --batch-id markets-1000-chunk-09 --mode quality --max-concurrent 3 --no-apollo --no-name-resolve --yes
```

**Terminal 10**
```powershell
.\.venv\Scripts\python.exe run_batch.py --input queries\chunks_1000\market_chunk_10.txt --batch-id markets-1000-chunk-10 --mode quality --max-concurrent 3 --no-apollo --no-name-resolve --yes
```

## Resume a chunk after crash

```powershell
cd D:\vendor-intel
$env:PYTHONPATH = ".\src"
.\.venv\Scripts\python.exe run_batch.py --batch-id markets-1000-chunk-01 --mode quality --max-concurrent 3 --no-apollo --no-name-resolve --yes
```
