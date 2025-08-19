import subprocess

for uuid in uuids:
    subprocess.run([
        "python", "report_generator.py",
        "--uuid", uuid,
        "--date_start", "2023-04-01",
        "--date_end", "2023-04-30"
    ])
