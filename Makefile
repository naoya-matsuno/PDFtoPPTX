default: dist/main.app
	open dist/main.app
dist/main.app: src/main.py
	poetry run pyinstaller src/main.py --onefile --noconsole --additional-hooks-dir hook -y