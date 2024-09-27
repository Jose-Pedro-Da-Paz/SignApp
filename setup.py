from setuptools import setup, find_packages

setup(
    name='Certificado de Visitação',
    version='1.0.0',
    description='Programa para automatizar a geração e envio dos Certificados de Visitação ',
    author='José Pedro Da Paz Ferreira',
    author_email='j.p.dapazferreira@hotmail.com',
    packages=find_packages(where='src'),
    package_dir={'': 'src'},
    install_requires=[
        'docx',  # Adicione outras dependências aqui
        'python-docx',
        'Pillow'
    ],
    entry_points={
        'console_scripts': [
            'meu_programa=meu_script:main2',  # Substitua "main" pela função principal do seu script
        ],
    },
    include_package_data=True,
)