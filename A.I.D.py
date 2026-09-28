import os, sys, argparse

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TESTS_DIR = os.path.join(BASE_DIR, "MenuInterfazUsuario")
if TESTS_DIR not in sys.path:
    sys.path.insert(0, TESTS_DIR)

def main():
    parser = argparse.ArgumentParser(description="Sistema Robot de Auditoría con interfaz")
    parser.add_argument("--gui", action="store_true", help="Lanzar GUI (Tkinter)")
    parser.add_argument("--web", action="store_true", help="Lanzar web (Flask)")
    args = parser.parse_args()

    # Por defecto o --gui
    from MenuInterfazUsuario.GuiMenu import run_gui
    run_gui()

if __name__ == "__main__":
    main()