from geopy.geocoders import Nominatim
from geopy.distance import vincenty
import openpyxl

print('Initializing...')
geolocator = Nominatim()
print('Loading .xlsx file...')
wb = openpyxl.load_workbook('travels.xlsx')

#creates new .xlsx file
print('Making report file...')
new_wb = openpyxl.Workbook()
new_wb_sheet = new_wb.get_active_sheet()
new_wb_sheet.title = 'Travel Counts'
new_wb_sheet['A1'] = 'Departure'
new_wb_sheet['B1'] = 'Arrival'
new_wb_sheet['C1'] = 'Distance'
new_wb_sheet['D1'] = 'Total Distance'
current_sheet = wb.get_sheet_by_name('Travel Origin and Destination')
print('Done')

i = 1
total_distance = 0
print('Gathering distances...')
print('--------')
while i <= 3489:
    try:
        depart_location = geolocator.geocode(current_sheet['D' + str(i)].value)
        depart_coordinates = (depart_location.latitude, depart_location.longitude)
        new_wb_sheet['A' + str(i+1)] = current_sheet['D' + str(i)].value
        
        arrival_location = geolocator.geocode(current_sheet['E' + str(i)].value)
        arrival_coordinates = (arrival_location.latitude, arrival_location.longitude)
        new_wb_sheet['B' + str(i+1)] = current_sheet['E' + str(i)].value
        
        distance = vincenty(depart_coordinates, arrival_coordinates).miles
        new_wb_sheet['C' + str(i+1)] = distance
        
        print('Trip ' + str(i) + ' is from ' + current_sheet['D' + str(i)].value + ' to ' + current_sheet['E' + str(i)].value + ': ' + str(distance) + ' miles')
        total_distance += distance
        new_wb_sheet['D' + str(i+1)] = total_distance
        print('Total distance is ' + str(total_distance) + ' miles')
        print('--------')
        new_wb.save('calculated_travels.xlsx')
        i += 1
        
    except AttributeError:
        print('!!!!!---Trip ' + str(i) + ' failed.---!!!!!')
        print('!!!!!---Do this manually---!!!!!')
        new_wb_sheet['A' + str(i+1)] = '~~~~~~Do this manually~~~~~'
        new_wb_sheet['B' + str(i+1)] = '~~~~~~Do this manually~~~~~'
        new_wb_sheet['C' + str(i+1)] = '~~~~~~Do this manually~~~~~'
        new_wb_sheet['D' + str(i+1)] = '~~~~~~Do this manually~~~~~'
        i += 1
    except socket.timeout:
        print('Somethign went wrong, trying again')
        i += 1
    except geopy.exc.GeocoderTimedOut:
        print('Something went wrong, trying again')
        i += 1
    except urllib.error.URLError:
        print('Something went wrong, trying again')
        i += 1
    except NameError:
        print('Something went wrong, trying again')
        i += 1
