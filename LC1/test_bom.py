import subprocess
import sys

# 模拟输入
inputs = [
    "",  # API key (empty)
    "2",  # choice
    "y",  # merge
    "BOM_complete.csv"  # filename
]

def run_with_inputs():
    process = subprocess.Popen([sys.executable, "main.py"],
                             stdin=subprocess.PIPE,
                             stdout=subprocess.PIPE,
                             stderr=subprocess.PIPE,
                             text=True)

    input_str = "\n".join(inputs) + "\n"
    stdout, stderr = process.communicate(input=input_str)

    print("STDOUT:")
    print(stdout)
    if stderr:
        print("STDERR:")
        print(stderr)

if __name__ == "__main__":
    run_with_inputs()