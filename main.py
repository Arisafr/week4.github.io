# website/views.py
from flask import Blueprint, render_template
from website import create_app

main = Blueprint('main', __name__)

@main.route('/create')
def create():
    app = create_app()

    if __name__ == '__main__':
        app.run(debug=True)
    return 'create function has been called'

@main.route('/home')
def home():
    return render_template('home.html')
