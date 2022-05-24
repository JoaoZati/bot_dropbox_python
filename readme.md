# EXCEL BOT

### Instalar requirements
```
conda env create --file requirements/botpython-conda.yml botpython-conda
ou
conda install --yes --file requirements/botpython-conda.yml

### Fazer update

conda env update --file requirements/botpython-conda.yml  --prune
ou
conda env update --name botpython-conda --file requirements/botpython-conda.yml --prune
```

### Exportar requirements
```
conda env export > requirements/botpython-conda.yml

### Instalar kivyMD

pip install pillow
pip install kivy
pip install https://github.com/kivymd/KivyMD/archive/master.zip

