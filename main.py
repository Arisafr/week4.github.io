from website import create_app


def start():
    # code for start function
    app = create_app()

    if __name__ == '__main__':
        app.run(debug=True)


