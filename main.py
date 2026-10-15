from gui import App, build_root


def main():
    root = build_root()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
