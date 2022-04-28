from geopy import geocoders


def geocode_coordinates(address, user_agent='Engineering_Services_BCC'):
    app = geocoders.Photon(user_agent=user_agent, proxies='165.225.226.22:10170')
    location = app.geocode(address)
    if location is not None:
        lat = location.latitude
        long = location.longitude
    else:
        lat = None
        long = None
    return lat, long, location

