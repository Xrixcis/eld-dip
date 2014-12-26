require_relative '../lib/exporter'

UNIT_QUERY = 'select 1'

DATA_QUERY = '
select
		region,
		pocet_jader,
		smer,
		lvl_jadraI,
		kod_from,
		o_from,
		from_x,
		from_y,
		sum_vyj,
  	sum_doj,
		lvl_jadraJ,
  	kod_to,
		o_to,
		to_x,
		to_y,
  	region_doj,
		tik,
		tki,
		tii,
		tjk,
		tkj,
		sum_tij,
		sum_tii,
  	curds_tij,
		curds_tji,
		min(pocet_jader, 3)
	from
		(select
			region,
			pocet_jader,
			(case
				when (sum_vyj > sum_doj) then 1
				when (sum_vyj < sum_doj) then 2
				else (case
					when (tki > tkj) then 1
					when (tki < tkj) then 2
					else 3 end)
				end) as smer,
			lvl_jadraI,
			kod_from,
			o_from,
			from_x,
			from_y,
			sum_vyj,
  		sum_doj,
			lvl_jadraJ,
  		kod_to,
			o_to,
			to_x,
			to_y,
  		region_doj,
			tik,
			tki,
			tii,
			tjk,
			tkj,
			sum_tij,
			sum_tii,
  		(sum_vyj/tkj + sum_vyj/tik + sum_doj/tjk + sum_doj/tki) as curds_tij,
			(sum_doj/tkj + sum_doj/tik + sum_vyj/tjk + sum_vyj/tki) as curds_tji
		from
		 	(select
				o1.region as region,
				(select count(*) from obce obc where obc.jadro = \'True\' and obc.region=o1.region) as pocet_jader,
				(case
					when t1.tki > 100000 then 1
	 				when t1.tki > 30000 then 2
					when t1.tki > 10000 then 3
					else 4 end) as lvl_jadraI,
  			kod_from,
				o1.nazev as o_from,
				o1.pos_x as from_x,
				o1.pos_y as from_y,
  			suma * 1.0 as sum_vyj,
  			coalesce((select suma from matice mx where mx.kod_from = m.kod_to and m.kod_from = mx.kod_to), 0) * 1.0 as sum_doj,
				(case
					when t2.tki > 100000 then 1
	 				when t2.tki > 30000 then 2
					when t2.tki > 10000 then 3
					else 4 end) as lvl_jadraJ,
  			kod_to,
				o2.nazev as o_to,
				o2.pos_x as to_x,
				o2.pos_y as to_y,
  			o2.region as region_doj,
				t1.tik as tik,
				t1.tki as tki,
				t1.tii as tii,
				t2.tik as tjk,
				t2.tki as tkj,
				(case when o1.region = o2.region then (select sum(suma) from matice mx join obce ox on (mx.kod_from = ox.kod) join obce ox2 on (mx.kod_to = ox2.kod) where ox.region = o1.region and ox2.region = o1.region)
					else (select sum(suma) from matice mx join obce ox on (mx.kod_from = ox.kod) join obce ox2 on (mx.kod_to = ox2.kod) where ox.skupina = o1.skupina and ox2.skupina = o1.skupina) end) as sum_tij,
				(case when o1.region = o2.region then (select sum(tii) from tiky tx join obce ox on (tx.kod = ox.kod) where ox.region = o1.region)
					else (select sum(tii) from tiky tx join obce ox on (tx.kod = ox.kod) where ox.skupina = o1.skupina) end) as sum_tii
			from
				matice m join tiky t1 on (m.kod_from = t1.kod)
				join tiky t2 on (m.kod_to = t2.kod)
				join obce o1 on (m.kod_from = o1.kod)
				join obce o2 on (m.kod_to = o2.kod)
			where
				o1.jadro = \'True\' and o2.jadro = \'True\' and o1.region=o2.region and 1=?
  			order by o1.region, o1.nazev, o2.nazev))'

def x(sym)
	HEADERS.each do |h|
		if h[0].eql? sym
			return h[2] + '$i'
		end
	end
end

HEADERS = [
		[:region, 'Region'],
		[:pocet_jader, 'Pocet_jader'],
		[:smer, 'Smer'],
		[:level_i, 'lvl_jadraI'],
		[:kod_obce_z, 'KODOB_VYJ'],
		[:nazev_obce_z, 'NAZOB_VYJ'],
		[:obec_z_x, 'FROM_X'],
		[:obec_z_y, 'FROM_Y'],
		[:sum_vyj, 'SumVyjDen'],
		[:sum_doj, 'SumDojDen'],
		[:level_j, 'lvl_jadraJ'],
		[:kod_obce_do, 'KODOB_DOJ'],
		[:nazev_obce_do, 'NAZOB_DOJ'],
		[:obec_do_x, 'TO_X'],
		[:obec_do_y, 'TO_Y'],
		[:region_do, 'REGION_DOJ'],
		[:tik, 'Tik'],
		[:tki, 'Tki'],
		[:tii, 'Tii'],
		[:tjk, 'Tjk'],
		[:tkj, 'Tkj'],
		[:sum_tij, 'SumOfTij'],
		[:sum_tii, 'SumOfTii'],
		[:curds_tij, 'CURDS_Tij'],
		[:curds_tji, 'CURDS_Tji'],
		[:kategorie_regionu, 'KATREG'],
		[:vztah, 'VZTAH'],
		[:curds_tij_n, 'CURDSPROC']
]

def excel_column(idx)
	a = (idx.div 26) - 1
	b = idx % 26
	if a < 0
		char = ''
	else
		char = ('A'.ord + a).chr
	end
	char << ('A'.ord + b).chr
	char
end

HEADERS.each_index do |idx|
	HEADERS[idx] << excel_column(idx)
end

FORMULAS = [
		"=(if(min(#{x :curds_tij}/#{x :curds_tji};#{x :curds_tji}/#{x :curds_tij}) &lt;= setup!$B$1 &amp; #{x :level_i}=#{x :level_j};1;
				if(min(#{x :curds_tij}/#{x :curds_tji};#{x :curds_tji}/#{x :curds_tij}) &gt; setup!$B$2 &amp; #{x :level_i} = #{x :level_j}; 2;
					if(min(#{x :curds_tij}/#{x :curds_tji};#{x :curds_tji}/#{x :curds_tij}) &lt;= setup!$B$1 &amp; #{x :level_i} &lt;&gt; #{x :level_j}; 3;
						4
		))))".gsub(/\s+/, ' '),
		"=(#{x :curds_tij}/setup!$B$3)*100"
]

exporter = Exporter.new 'all_per_region_sum.xlsx', 'all_per_region', HEADERS.map {|h| h[1]}, FORMULAS, {name: 'setup', data: [['Koeficient <=:', '0.5'], ['Koeficient >:', '0.5'], ['Max curds_tij:', '0']]}
exporter.run('database.sqlite', UNIT_QUERY, DATA_QUERY)
