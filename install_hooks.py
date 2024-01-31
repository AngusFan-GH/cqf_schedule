import os
import shutil


def install_hook(script_name):
    # install the hook
    hooks_dir = os.path.join('.git', 'hooks')
    custom_hooks_dir = 'custom_hooks'
    script_path = os.path.join(custom_hooks_dir, script_name)
    target_path = os.path.join(hooks_dir, script_name)

    if os.path.isfile(script_path):
        shutil.copy(script_path, target_path)
        print(f"{script_name} hook installed.")
    else:
        print(f"Error: {script_path} does not exist.")


if __name__ == "__main__":
    install_hook('post-commit')
