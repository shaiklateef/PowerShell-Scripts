(GWMI win32_pnpentity -Filter "caption like '%Intel(R) Wireless Bluetooth(R)%'").disable()


#GWMI win32_pnpentity -Filter "caption like '%Bluetooth%'" | Select caption


#(GWMI win32_pnpentity -Filter "caption like '%Intel(R) Wireless Bluetooth(R)%'").enable()