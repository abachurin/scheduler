from setuptools import setup, find_packages

with open('requirements.txt') as fh:
    install_requires = fh.read().split('\n')

setup(name='Scheduler',
      version="1.0",
      description='Task Scheduler',
      author='Alex Bachurin',
      author_email='bachurin.alex@gmail.com',
      python_requires='>=3.9',
      packages=find_packages(),
      include_package_data=True,
      data_files=[('', ['requirements.txt'])],
      install_requires=install_requires[:-1]
      )
