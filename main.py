import subprocess

def run_script(script_name):
    """Run a Python script using subprocess."""
    result = subprocess.run(['python', script_name], capture_output=True, text=True)
    return result.returncode, result.stdout, result.stderr

def main():
    
    return_code, stdout, stderr = run_script('sort.py')
    if return_code == 0:
        print("sort.py executed successfully.")
        print(stdout)
    else:
        print("Error executing sort.py:")
        print(stderr)
        return

    
    return_code, stdout, stderr = run_script('sort2.py')
    if return_code == 0:
        print("sort2.py executed successfully.")
        print(stdout)
    else:
        print("Error executing sort2.py:")
        print(stderr)
        return
    
    return_code, stdout, stderr = run_script('sort3.py')
    if return_code == 0:
        print("sort3.py executed successfully.")
        print(stdout)
    else:
        print("Error executing sort3.py:")
        print(stderr)
        return

    print("Three scripts executed successfully. Final output is sorted_output2.xlsx.")

if __name__ == "__main__":
    main()
