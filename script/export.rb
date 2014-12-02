require_relative '../lib/exporter'

UNIT_QUERY = 'select distinct region from obce'

DATA_QUERY = 'select
	o1.region as region,
  kod_from,
	o1.nazev as o_from,
  suma,
  coalesce((select suma from matice mx where mx.kod_from = m.kod_to and m.kod_from = mx.kod_to), 0) as suma2,
  kod_to,
	o2.nazev as o_to,
	t1.tik as tik,
	t1.tki as tki,
	t1.tii as tii,
	t2.tik as tjk,
	t2.tki as tkj,
	(select sum(suma) from matice mx join obce ox on (mx.kod_from = ox.kod) join obce ox2 on (mx.kod_to = ox2.kod) where ox.region = o1.region and ox2.region = o1.region) as sum_tij,
	(select sum(tii) from tiky tx join obce ox on (tx.kod = ox.kod) where ox.region = o1.region) as sum_tii
from
	matice m join tiky t1 on (m.kod_from = t1.kod)
	join tiky t2 on (m.kod_to = t2.kod)
	join obce o1 on (m.kod_from = o1.kod)
	join obce o2 on (m.kod_to = o2.kod)
where
	o1.region = o2.region and o1.region = ?
  order by o1.region, o1.nazev, o2.nazev'

HEADERS = ['Region', 'KODOB_VYJ', 'NAZOB_VYJ', 'SumOfVYJ_DENNE', 'SumOfDOJ_DENNE', 'KODOB_DOJ', 'NAZOB_DOJ', 'Tik', 'Tki', 'Tii', 'Tjk', 'Tkj', 'SumOfTij', 'SumOfTii', 'poměr Tij(sumaTij+sumaTii)', 'poměr Tji/(sumaTij+sumaTii)', 'SMART_Tij', 'SMART_Tji', 'INTRAMAX_Tij', 'INTRA_Tji', 'CURDS_Tij', 'CURDS_Tji']
FORMULAS = ['=D$i/(M$i+N$i)', '=E$i/(M$i+N$i)', '=(D$i*D$i)/(L$i*H$i)+(E$i*E$i)/(K$i*I$i)', '=(E$i*E$i)/(L$i*H$i)+(D$i*D$i)/(K$i*I$i)', '=D$i/(L$i*H$i)+E$i/(K$i*I$i)', '=E$i/(L$i*H$i)+D$i/(K$i*I$i)', '=D$i/L$i+D$i/H$i+E$i/K$i+E$i/I$i', '=E$i/L$i+E$i/H$i+D$i/K$i+D$i/I$i']

exporter = Exporter.new 'regiony_sum.xlsx', 'regiony', HEADERS, FORMULAS
exporter.run 'database.sqlite', UNIT_QUERY, DATA_QUERY