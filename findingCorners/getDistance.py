# First option(less precise)

from geopy.geocoders import Nominatim
from geopy import distance

geocoder =  Nominatim(user_agent="maps")
location1 = "Los Angeles, California, USA"
location2 = "San Francisco, California, USA"

coordinates1 = geocoder.geocode(location1)
coordinates2 = geocoder.geocode(location2)
lat1,long1 = (coordinates1.latitude),(coordinates1.longitude)
lat2,long2 = (coordinates2.latitude),(coordinates2.longitude)

place1 = (lat1,long1)
place2 = (lat2,long2)

print(distance.distance(place1,place2).miles)

#-------------------------------------------------------------------

# Second option(less convinient)

from geopy.distance import great_circle

location1 = (34.056920236361215, -118.2429441850822)
location2 = (37.77567677707561, -122.42508975426993)

print(great_circle(location1, location2).miles)