from geopy import geocoders
import time
google_geo_code_key = 'AIzaSyDgFypRMb-gnE9eaFjiWjcdc6T4JpjGUAo'

def geocode_coordinates(address, user_agent='Engineering_Services_BCC', api='osm'):
    lat_ = None
    long_ = None
    location_ = None
    if api.lower() == 'osm':
        app = geocoders.Photon(user_agent=user_agent, proxies='165.225.226.22:10170')
        time.sleep(2)
    elif api.lower() == 'here':
        app = geocoders.HereV7(user_agent=user_agent, proxies='165.225.226.22:10170', apikey='hROuZ5fSMweHJUgssiq6oehaPsd6u8-qMeF6CGN-SOQ')
        time.sleep(2)
    elif api.lower() == 'arcgis':
        app = geocoders.ArcGIS(user_agent=user_agent, proxies='165.225.226.22:10170')
        time.sleep(2)
    elif api.lower() == 'google':
        app = geocoders.GoogleV3(user_agent=user_agent, proxies='165.225.226.22:10170', api_key=google_geo_code_key)
        time.sleep(2)
    print(address, api)
    location_ = app.geocode(address)

    if location_ is not None:
        lat_ = location_.latitude
        long_ = location_.longitude
    else:
        lat_ = None
        long_ = None
    return lat_, long_, location_

location = 'Ann Street Between Roma St and Edward St BRISBANE_CITY'
lat, long, location_out = geocode_coordinates(location, api='osm')
print('osm', lat, long, location_out)

lat, long, location_out = geocode_coordinates(location, api='here')
print('here', lat, long, location_out)

lat, long, location_out = geocode_coordinates(location, api='arcgis')
print('arcgis', lat, long, location_out)

lat, long, location_out = geocode_coordinates(location, api='google')
print('google', lat, long, location_out)