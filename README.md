# Reporte en Excel
## Archivos
- Imágenes png: son todas las imágenes que se generan en el código "gen_excel.py". Sirven para poder pegarlas en las hojas de trabajo.
- 2017_prediction.csv: fichero csv que contiene la predicción de ingredientes por semana para 2017. Este fichero ha sido importado directamente, no se genera en este código
- df_merged.csv: dataframe que resulta de la combinación de orders.csv y de order_details.csv. Este fichero ha sido importado directamente, no se genera en este código.
- pizzas.csv: fichero csv que contiene información acerca de cada modelo de pizzas.

## Descripción del script
En este script, primero se leen los csv correspondientes y se hacen algunas modificaciones de algunos de ellos para que, a la hora de crear las gráficas, sea mucho más cómodo. Por otro lado, se utiliza la librería openpyxl para construir el fichero .xlsx, mediante la cual tanto se generan gráficas directamente en excel como se pegan otras guardadas previamente. Hojas de trabajo:
- Reporte ejecutivo: contiene gráficas relacionadas con los ingresos.
- Reporte de ingredientes: contiene las gráficas que representan cuántos ingredientes hacen falta por semana.
- Reporte de pedidos: Contiene gráficas que expresan información acerca de los pedidos, por ejemplo, cuántas pizzas se han vendido por mes,
