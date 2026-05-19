import http.client

host = 'jooble.org'
key = 'b22be733-605d-4a4f-b781-bc3aaa8bec6c'

connection = http.client.HTTPConnection(host)
#request headers
headers = {"Content-type": "application/json"}
#json query
body = '{ "keywords": "senior software engineer", "location": "New York"}'
connection.request('POST','/api/' + key, body, headers)
    response = connection.getresponse()
    print(response.status, response.reason)
    print(response.read())
except Exception as e:
    print(e)
finally:
    connection.close()