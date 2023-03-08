from flask import Flask

def create_app():
    app = Flask(__name__)
    app.config['SECRET_KEY'] = 'secretkey'
    
    from .views import views
    app.register_blueprint(views,url_prefix='/')

    from .pros import pros
    app.register_blueprint(pros,url_prefix='/')

    from ..main import main
    app.register_blueprint(main,url_prefix='/')


    return app

