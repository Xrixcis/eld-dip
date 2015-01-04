require_relative '../lib/simple_exporter'

exporter = SimpleExporter.new 'jadra.xlsx', %w(nazev rozloha obyvatel Tki level region KATREG), []
exporter.run('database2.sqlite',
             'select nazev,
                     area,
                     obyvatel,
                     tki,
                     (case
                                when tki > 100000 then 1
                                when tki > 30000 then 2
                                when tki > 10000 then 3
                                else 4 end) as level,
                     region,
                     min((select count(*) from obce obc where obc.jadro = \'True\' and obc.region=o.region), 3) as katreg
             from obce o join tiky using(kod)
             where jadro=\'True\'
             order by region, nazev')

exporter = SimpleExporter.new 'katreg.xlsx', %w(katreg regionu obyvatel rozloha Tki jader), []
exporter.run('database2.sqlite',
             'select katreg,
                count(distinct region) as regionu,
                sum(obyvatel) as obyvatel,
                sum(area) as rozloha,
                sum(tki) as tki,
                sum((case when jadro = \'True\' then 1 else 0 end)) as jader
              from (select
                     area,
                     obyvatel,
                     tki,
                     region,
                     min((select count(*) from obce obc where obc.jadro = \'True\' and obc.region=o.region), 3) as katreg,
                     jadro
             from obce o join tiky using(kod)) group by katreg')


exporter = SimpleExporter.new 'regiony.xlsx', %w(region obyvatel rozloha jader KATREG Tki), []
exporter.run('database2.sqlite',
             'select region,
                sum(obyvatel) as obyvatel,
                sum(area) as rozloha,
                sum((case when jadro = \'True\' then 1 else 0 end)) as jader,
                katreg,
                sum(tki) as tki
              from (select
                     area,
                     obyvatel,
                     tki,
                     region,
                     min((select count(*) from obce obc where obc.jadro = \'True\' and obc.region=o.region), 3) as katreg,
                     jadro
             from obce o join tiky using(kod)) group by region')