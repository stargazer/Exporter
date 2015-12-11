from setuptools import setup

setup(
    name='exporter',
    packages=('exporter',),
    author='C. Paschalides',
    author_email='already.late@gmail.com',
    description='Export python data structures to a few file formats.',
    url='https://github.com/stargazer/exporter/',
    license='WTFPL',
    version=0.1,
    install_requires=('tablib',),
    zip_safe=False,
    classifiers=(
        'Programming Language :: Python',
        'Operating System :: OS Independent',
        'Topic :: Internet :: WWW/HTTP',
        'Intended Audience :: Developers',
        'License :: Freely Distributable'
    ),
)
