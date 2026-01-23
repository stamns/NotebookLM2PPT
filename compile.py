import toml
import os
import argparse

# pyinstaller --clean -F -w -n notebooklm2ppt_{version} --optimize=2 --collect-all spire.presentation main.py 

if __name__ == "__main__":
    # 读取toml中的版本号
    # 执行编译命令

    parser = argparse.ArgumentParser(description="Compile notebooklm2ppt into a standalone executable.")
    parser.add_argument(
        "--as_dir",
        action="store_true",
        help="Compile as a directory instead of a single file executable.",
    )
    args = parser.parse_args()

    
    with open("pyproject.toml", "r", encoding="utf-8") as f:
        pyproject_data = toml.load(f)


    version = pyproject_data["project"]["version"]
    output_name = f"notebooklm2ppt-{version}"
    print(f"编译版本: {output_name}")
    os.system('del *.spec')
    
    if not args.as_dir:
        flags = "-F"
    else:
        flags = "-D"
    command = f'pyinstaller --clean {flags} -w -n {output_name} --optimize=2 --collect-all spire.presentation main.py'
    os.system(command)
