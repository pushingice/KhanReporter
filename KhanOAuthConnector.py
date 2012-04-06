# -*- coding: utf-8 *-*

import cgi
import SocketServer
import SimpleHTTPServer

from test_oauth_client import TestOAuthClient
from oauth import OAuthToken
import time
import json
import os.path
import sys
import LoadStudents

REQUEST_TOKEN = None


class KhanOAuth():

    CONSUMER_KEY = ""
    CONSUMER_SECRET = ""
    SERVER_URL = "http://www.khanacademy.org"
    ACCESS_TOKEN = None

    def initConnection(self):

        oauthPath = os.path.join("private", "OAuth.txt")
        file = open(oauthPath).readlines()
        self.CONSUMER_KEY = file[0].split()[1]
        self.CONSUMER_SECRET = file[1].split()[1]
        self.getTokens()

    def create_callback_server(self):
        class CallbackHandler(SimpleHTTPServer.SimpleHTTPRequestHandler):
            def do_GET(self):
                global REQUEST_TOKEN
                params = cgi.parse_qs(self.path.split('?', 1)[1], \
                    keep_blank_values=False)
                REQUEST_TOKEN = OAuthToken(params['oauth_token'][0], \
                    params['oauth_token_secret'][0])
                REQUEST_TOKEN.set_verifier(params['oauth_verifier'][0])

                self.send_response(200)
                self.send_header('Content-Type', 'text/plain')
                self.end_headers()
                self.wfile.write(\
                    'OAuth request token fetched; you can close this window.')

            def log_request(self, code='-', size='-'):
                pass

        server = SocketServer.TCPServer(('127.0.0.1', 0), CallbackHandler)
        return server

    def get_request_token(self):
        server = self.create_callback_server()

        client = TestOAuthClient(self.SERVER_URL, self.CONSUMER_KEY, \
            self.CONSUMER_SECRET)
        client.start_fetch_request_token(\
            'http://127.0.0.1:%d/' % server.server_address[1])

        server.handle_request()
        # REQUEST_TOKEN has now been set
        server.server_close()

    def get_access_token(self):

        client = TestOAuthClient(self.SERVER_URL, self.CONSUMER_KEY, \
            self.CONSUMER_SECRET)
        self.ACCESS_TOKEN = client.fetch_access_token(REQUEST_TOKEN)

    def get_api_resource(self, resourceUrl):

        # Example URLs
        #/api/v1/user/exercises/exponents_1?email=Khan.User@gmail.com
        #/api/v1/user?email=Khan.User@gmail.com
        #/api/v1/exercises

        client = TestOAuthClient(self.SERVER_URL, self.CONSUMER_KEY, \
            self.CONSUMER_SECRET)
        start = time.time()
        response = client.access_resource(resourceUrl, self.ACCESS_TOKEN)
        end = time.time()

        #print "\nTime: %ss\n" % (end - start)
        return response

    def getTokens(self):

        self.get_request_token()
        if not REQUEST_TOKEN:
            print "Did not get request token."
            return

        self.get_access_token()
        if not self.ACCESS_TOKEN:
            print "Did not get access token."
            return

if __name__ == "__main__":

    KOA = KhanOAuth()
    KOA.initConnection()
    studentMap = LoadStudents.returnStudentMap();
    response = []

    for studentUUID in studentMap:
        try:
            response = KOA.get_api_resource("/api/v1/user?email=" + studentUUID)
        except EOFError:
            print
        except Exception, e:
            print "Error: %s" % e
        
        print("---")
        if (response.strip() != 'null'):
            jsonObject = json.loads(response)
            profList = jsonObject.get("all_proficient_exercises")
            nickname = jsonObject.get("nickname")
            print nickname
            for p in profList:
                exerciseResponse = KOA.get_api_resource \
                    ("/api/v1/user/exercises/" + p + "?email=" + studentUUID)
                if (exerciseResponse.strip() != 'null'):
                    exerciseJSON = json.loads(exerciseResponse)
                    proficientDate = exerciseJSON.get("proficient_date")
                    sys.stdout.write(p + " ")
                    if (proficientDate): sys.stdout.write(proficientDate)
                    else: sys.stdout.write("null")
                    sys.stdout.write('\n')
            
    
