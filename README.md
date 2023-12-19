Program for comparing product prices in K-rauta, Puuilo, and Prisma.
V.1:
  Can search products based on EAN/product code, there's essentially no error handling, the code is slow, and if the product can't be found there's the logic for that only with the Prisma website.

V.2
  Reads EAN codes from an Excel file, searches products from Puuilo, Prisma, K-rauta, and Tokmanni, and writes each store's prices (except for Prisma which is only used for title and manufacturer info) to the same   
  excel in their respective columns.
  Changed the structure of the program so that functions are grouped into classes, implemented some parallel processing with concurrent.futures and lambda, switched to local driver session (that is also ended omg), 
  and changed the sleep-time to comply the websites' robots.txt.
