#!/usr/bin/env python3
"""Decode a persisted Google Drive MCP download result into a real file.
The Drive connector's download_file_content returns base64 JSON; for files above a few hundred KB the harness saves
the result under ~/.claude/projects/<project>/<session>/tool-results/mcp-Google_Drive-download_file_content-*.txt.
Usage: decode_drive_result.py <tool-result.txt> <out_dir>   (writes <out_dir>/<title>)"""
import sys, json, base64, os
src, out_dir = sys.argv[1], sys.argv[2]
j = json.load(open(src)); os.makedirs(out_dir, exist_ok=True)
path = os.path.join(out_dir, j['title']); open(path, 'wb').write(base64.b64decode(j['content']))
print(path, os.path.getsize(path), 'bytes', j.get('mimeType'))
