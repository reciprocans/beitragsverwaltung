from werkzeug.middleware.dispatcher import DispatcherMiddleware
from werkzeug.wrappers import Response
from app import app as vereins_app
from vorstandsprotokolle.app import app as protokoll_app

application = DispatcherMiddleware(vereins_app, {'/protokolle': protokoll_app})
