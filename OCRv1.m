%Intento de programa que escanee el recibo de cfe 
%para que haga todo lo que yo no quiero c: 


%DEFINICION DEL RECIBO A LEER 
%buscar ciclar esto pls CON VECTOR DE IMAGENES
I = imread('1708.jpg');

%DEFINICION DEL ARCHIVO DE EXCEL
excelfile = 'OCR.xlsx';

%DEFINIR PUNTOS DE INTERES mediante interaccion del usuario 
%comentar una vez obtenidos los valores
%figure;
%imshow(I);
%TEMPROI = round(getPosition(imrect))
%close
kWhROI = [649   578    77    25];
kWhocr = ocr(I, kWhROI,'CharacterSet','0123456789,');
kWh = str2double(kWhocr.Text);

%Escribir el valor kWh obtenido en el archivo de excel 
xlswrite(excelfile,kWh,'Hoja1','C7');

%=======================================================
kWROI = [664   604    46    25];
kWocr = ocr(I, kWROI,'CharacterSet','0123456789,');
kW = str2double(kWocr.Words);
xlswrite(excelfile,kW,'Hoja1','D7');

%=======================================================
kVAhrROI = [653   629    64    31];
kVAhrocr = ocr(I, kVAhrROI,'CharacterSet','0123456789,');
kVAhr = str2double(kVAhrocr.Words);
xlswrite(excelfile,kVAhr,'Hoja1','E7');
%=======================================================
FPROI = [650        1031          89          36];
%Para el factor de potencia, al tener punto decimal se agrega en el
%CharacterSet y se quita la coma
FPocr = ocr(I, FPROI,'CharacterSet','0123456789.');
FP = str2double(FPocr.Words);
xlswrite(excelfile,FP,'Hoja1','F7');
%========================================================
DemandaFacturableROI = [289        1003         115          26];
DemandaFacturableocr = ocr(I, DemandaFacturableROI);
DemandaFacturable = str2double(DemandaFacturableocr.Words);
xlswrite(excelfile,DemandaFacturable,'Hoja1','G7');
%========================================================
TarifaEnergeticaROI = [670   813   132    44];
TarifaEnergeticaocr = ocr(I, TarifaEnergeticaROI,'CharacterSet','0123456789.,');
TarifaEnergetica = str2double(TarifaEnergeticaocr.Words);
xlswrite(excelfile,TarifaEnergetica,'Hoja1','H7');
%========================================================
DemandaFacturableROI = [547        1001          97          29];
DemandaFacturableocr = ocr(I, DemandaFacturableROI,'TextLayout','Word','CharacterSet','0123456789.,');
DemandaFacturable = str2double(DemandaFacturableocr.Words);
xlswrite(excelfile,DemandaFacturable,'Hoja1','I7');
%========================================================

%ARREGLAR LA COMPROBACION DEL SIGNO POSITIVO O NEGATIVO%
%SOLUCIONADO.

TarifaFPROI = [1518        1456          82          29];
TarifaFPocr = ocr(I, TarifaFPROI,'TextLayout','Word','CharacterSet','0123456789.,-');
TarifaFP = str2double(TarifaFPocr.Words);
%xlswrite(excelfile,TarifaFP,'Hoja1','J7');
%stringtarifa
xlswrite(excelfile,TarifaFPocr.Words,'Hoja1','J7');
%========================================================
IVAROI = [1502        1514          94          26];
IVAocr = ocr(I, IVAROI,'TextLayout','Word','CharacterSet','0123456789.,');
IVA = str2double(IVAocr.Words);
xlswrite(excelfile,IVA,'Hoja1','K7');


%========================================================
AlumbradoROI = [1507        1566          94          31];
Alumbradoocr = ocr(I, AlumbradoROI,'TextLayout','Word','CharacterSet','0123456789.,');
Alumbrado = str2double(Alumbradoocr.Words);
xlswrite(excelfile,Alumbrado,'Hoja1','L7');
%=========================================================
TotalFacturadoROI = [1470        1621         128          27];
TotalFacturadoocr = ocr(I, TotalFacturadoROI,'TextLayout','Word','CharacterSet','0123456789.,$');
%TotalFacturado = str2double(TotalFacturadoocr.Words);
xlswrite(excelfile,TotalFacturadoocr.Words,'Hoja1','O7');

%========================================================
%ABRIR EXCEL PARA VERIFICAR 
winopen('OCR.xlsx')
%========================================================
% figure;
% imshow(I);
% TEMPROI = round(getPosition(imrect))
% close