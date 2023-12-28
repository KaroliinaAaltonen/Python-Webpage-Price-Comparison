Program for comparing product prices in K-rauta, Puuilo, Tokmanni, Bauhaus, and Prisma.

V.1:
  Can search products based on EAN/product code, there's essentially no error handling, the code is slow, and if the product can't be found there's the logic for that only with the Prisma website.

V.2
  Reads EAN codes from an Excel file, searches products from Puuilo, Prisma, K-rauta, and Tokmanni, and writes each store's prices (except for Prisma which is only used for title and manufacturer info) to the same   
  excel in their respective columns.
  Changed the structure of the program so that functions are grouped into classes, implemented some parallel processing with concurrent.futures and lambda, switched to local driver session (that is also ended omg), 
  and changed the sleep-time to comply the websites' robots.txt.

V.3
  Also handles Bauhaus data, product links are now directly to product page instead of search page. Sound notification when program is done executing. K-Rauta          scraping doesn't have function to check if the product can be found so currently the code probably crashes if some product doesn't exist on their website.

V.4
Scrapper classes are uniform. There is error handling in each function (just basic try-except Exception). 
