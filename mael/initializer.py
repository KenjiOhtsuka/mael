import os
import re
import shutil


class Initializer:
    def __init__(self, dir_path):
        self.dir_path = dir_path

    def initialize(self):
        dir_path = self.dir_path
        print(f'Initialize {dir_path}.')
        while True:
            templates = {
                0: 'Normal',
                1: 'Test case'
            }
            while True:
                template_answer = input(
                    "\n".join([
                        "Which template do you use?\n",
                        *[f'{k}: {v}' for k, v in templates.items()],
                        '\nType number: '
                    ])
                )
                if not re.match(r'\d', template_answer):
                    continue
                template_answer = int(template_answer)
                if template_answer in templates:
                    break
            break

        file_path = os.path.abspath(__file__)
        template_path = os.path.join(
            os.path.dirname(file_path),
            'templates',
            templates[template_answer].lower().replace(' ', '_')
        )

        if not os.path.exists(dir_path):
            os.mkdir(dir_path)
        shutil.copytree(template_path, dir_path, dirs_exist_ok=True)
