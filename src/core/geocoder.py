from pathlib import Path
import csv
import logging
from geopy.geocoders import Nominatim

class Geocoder:
    """Handles geocoding functionality"""
    
    def __init__(self, cache_file):
        self.geolocator = Nominatim(user_agent="pole_mapper")
        self.cache_file = cache_file
        self.cache = self.load_cache()
    
    def load_cache(self):
        cache = {}
        if self.cache_file.exists():
            try:
                with open(self.cache_file, 'r', encoding='utf-8') as f:
                    reader = csv.DictReader(f)
                    for row in reader:
                        key = f"{row['latitude']},{row['longitude']}"
                        cache[key] = row['address']
            except:
                pass
        return cache
    
    def reverse(self, lat, lon):
        try:
            lat, lon = float(lat), float(lon)
            key = f"{round(lat, 7)},{round(lon, 7)}"
            
            if key in self.cache:
                return self.cache[key]
            
            location = self.geolocator.reverse(f"{lat}, {lon}", timeout=10)
            address = self.format_address(location.raw.get('address', {})) if location else ''
            
            self.save_to_cache(key, address)
            return address
        except:
            return ''
    
    def format_address(self, addr):
        """Format address in the desired format: '14405 Cameo Avenue, Minnesota 55068, US'"""
        parts = []
        
        # Add house number and road
        if addr.get('house_number') and addr.get('road'):
            parts.append(f"{addr['house_number']} {addr['road']}")
        elif addr.get('road'):
            parts.append(addr['road'])
        
        # Add state and postcode
        if addr.get('state'):
            state_postcode = addr['state']
            if addr.get('postcode'):
                state_postcode += f" {addr['postcode']}"
            parts.append(state_postcode)
        
        # Add country
        if addr.get('country_code'):
            parts.append(addr['country_code'].upper())
        
        return ', '.join(parts)
    
    def save_to_cache(self, key, address):
        try:
            lat, lon = key.split(',')
            with open(self.cache_file, 'a', newline='', encoding='utf-8') as f:
                writer = csv.DictWriter(f, fieldnames=["latitude", "longitude", "address"])
                if self.cache_file.stat().st_size == 0:
                    writer.writeheader()
                writer.writerow({"latitude": lat, "longitude": lon, "address": address})
            self.cache[key] = address
        except:
            pass