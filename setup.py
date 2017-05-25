import ast
import os.path

from setuptools import setup


def get_version():
    module_path = os.path.join(os.path.dirname(__file__), 'dodotable_xlsx.py')
    module_file = open(module_path)
    try:
        module_code = module_file.read()
    finally:
        module_file.close()
    tree = ast.parse(module_code, module_path)
    for node in ast.iter_child_nodes(tree):
        if not isinstance(node, ast.Assign) or len(node.targets) != 1:
            continue
        target, = node.targets
        if isinstance(target, ast.Name) and target.id == '__version__':
            value = node.value
            if isinstance(value, ast.Str):
                return value.s
            raise ValueError('__version__ is not defined as a string literal')
    raise ValueError('could not find __version__')


def readme():
    path = os.path.join(os.path.dirname(__file__), 'README.rst')
    try:
        with open(path) as f:
            return f.read()
    except IOError:
        pass


setup(
    name='dodotable-xlsx',
    version=get_version(),
    description='Excel (.xlsx) exporter for dodotable',
    long_description=readme(),
    url='https://github.com/spoqa/dodotable-xlsx',
    author='Hong Minhee',
    author_email='hongminhee' '@' 'spoqa.com',
    license='MIT license',
    py_modules=['dodotable_xlsx'],
    install_requires=[
        'dodotable >= 0.4.0, < 1.0.0',
        'MarkupSafe',
        'XlsxWriter >= 0.9.6, < 1.0.0',
    ],
)
