# encoding: utf-8
module XLSSestavy
  module ExcelTabulky

    # vytvoří list, zapíše jej do @ws a vrátí jej. mezi tím případně provede předaný blok
    def vytvor_list(nazev)
      @ws = @wb.add_worksheet nazev
      yield @ws if block_given?
      @ws
    end

    # roztahuje se na (defaultně) 24 sloupců v prvním řádku
    def sestava_nadpis(text, roztahnout=24, radek=0)
      @ws.merge_range radek, 0, radek, roztahnout-1, text, get_format(:sestava_nadpis)
      @ws.set_row radek, XLSSestavy.row_cm_to_p(1)
    end

    # pozice ve standartních excel souřadnicích 'A2', 'B3:B5' (spojení buňek)
    # vyska znamená výška daného řádku v cm
    def sestava_napdis2(text, pozice, vyska = 0.7)
      zapis pozice, text, get_format(:sestava_nadpis2)
      @ws.set_row ciselne_souradnice(pozice)[0][0], XLSSestavy.row_cm_to_p(vyska)
    end

    #pozice ve standartních excel souřadnicích 'A2', 'B3:B5' (spojení buňek)
    def sestava_cas_vytvoreni(pozice='A2')
      cas = "začátek zpracování: #{l Time.now}"
      zapis pozice, cas, get_format(:default)
    end

    # pozice bunkdy jsou ve standartních souřadnicích ( 'A3')
    def zapis(bunky, hodnota, format=nil)
      format = get_format unless format
      souradnice = XLSSestavy.ciselne_souradnice bunky
      if souradnice.length==1
        @ws.write souradnice[0][0], souradnice[0][1], hodnota, format
      else
        @ws.merge_range souradnice[0][0], souradnice[0][1], souradnice[1][0], souradnice[1][1], hodnota, format
      end
    end

    def zapis_radu(prvni_bunka, hodnoty, format=nil)
      format = get_format unless format
      souradnice = XLSSestavy.ciselne_souradnice(prvni_bunka).first
      hodnoty.each do |h|
        @ws.write souradnice[0], souradnice[1], h, format
        souradnice[1] += 1
      end
    end

    def zapis_sloupec(prvni_bunka, hodnoty, format=nil)
      format = get_format unless format
      souradnice = XLSSestavy.ciselne_souradnice(prvni_bunka).first
      hodnoty.each do |h|
        @ws.write souradnice[0], souradnice[1], h, format
        souradnice[0] += 1
      end
    end

    #vypsání dat tabulky, hlaviček a případných součtových řádků
    # vrací počet, kolik řádků bylo vypsáno
    # objekty = <Array<Object>>  /  <ActiveRecord::Relation>
    # sloupce: <Array<Sloupec>> / <RadaSloupcu> -pole sloupců
    #--- args ---
    # vyska_zahlavi: <Numerical>  -hodnota výšky prvního řádku v cm
    # ukotvit_zahlavi: <True> / <Nil>  -pokud má být za hlavičkou ukotveno(o jeden řádek níže při použití součtových řádků :nad)
    # soucty: :nad, :pod, :nad_pod, :prazdne, nil
    # format_zahlavi: <Symbol>  -formát hlavičky (bude upraven formátem sloupce)
    # format_dat: <Symbol>  -formát řádků dat (bude upraven formátem sloupce)
    def vypis_tabulku(pozice, objekty, sloupce, args={})
      y, x = XLSSestavy.ciselne_souradnice(pozice).first
      sloupce = sloupce.sloupce if sloupce.kind_of? RadaSloupcu
      soucty = args[:soucty]
      soucty = :prazdne if objekty.length==0 && soucty

      dy = 0 # posun v řádcích
      #hlavičky
      format_zahlavi = get_format args[:format_zahlavi]||:zahlavi
      formaty = sloupce.map do |s|
        next format_zahlavi unless s.format.class==Hash
        format_hash = s.format.clone
        format_hash[:num_format] = nil
        alter_format format_zahlavi, format_hash
      end
      sloupce.each_with_index do |s, i|
        x_sloupce = x + i
        @ws.write y, x_sloupce, s.zahlavi, formaty[i]
        @ws.set_column x_sloupce, x_sloupce, XLSSestavy.col_cm_to_p(s.sirka)
      end
      @ws.set_row y, XLSSestavy.row_cm_to_p(args[:vyska_zahlavi]||1.3)
      dy += 1
      #součty :nad
      y_soucty = y+dy+1
      if soucty==:nad || soucty==:nad_pod
        y_soucty += 1
        vypis_souctovy_radek y+dy, x, sloupce, (y_soucty..y_soucty+objekty.length-1)
        dy += 1
      end
      #ukotvit
      @ws.freeze_panes y+dy, 0 if args[:ukotvit_zahlavi]
      #vypsani samotnych dat
      format_dat = get_format args[:format_dat]||:data
      formaty = sloupce.map do |s|
        next format_dat unless s.format.class==Hash
        alter_format format_dat, s.format
      end
      if objekty.class==ActiveRecord::Relation
        objekty.find_in_batches batch_size: 100 do |batch|
          batch.each do |objekt|
            sloupce.each_with_index do |s, i|
              hodnota = XLSSestavy.douprav_hodnotu_bunky s.hodnota_pro(objekt)
              if s.num_format==:cas || s.num_format==:datum
                @ws.write_date_time y+dy, x+i, hodnota, formaty[i]
              else
                @ws.write y+dy, x+i, hodnota, formaty[i]
              end
            end
            dy += 1
          end
        end
      else
        objekty.each do |objekt|
          sloupce.each_with_index do |s, i|
            hodnota = XLSSestavy.douprav_hodnotu_bunky s.hodnota_pro(objekt)
            if s.num_format==:cas || s.num_format==:datum
              @ws.write_date_time y+dy, x+i, hodnota, formaty[i]
            else
              @ws.write y+dy, x+i, hodnota, formaty[i]
            end
          end
          dy += 1
        end
      end
      #součty :pod
      if soucty==:pod || soucty==:nad_pod
        vypis_souctovy_radek y+dy, x, sloupce, (y_soucty..y_soucty+objekty.length-1)
        dy += 1
      end
      #případně, když nejsou žádné objekty jenom prázdný součtový řádek
      if soucty==:prazdne
        format = get_format :souctovy_radek
        sloupce.length.times{|i| @ws.write y+dy, x+i, '', format }
        dy += 1
      end
      dy
    end

    def vypis_souctovy_radek(y, x, sloupce, rozsah)
      format = get_format :souctovy_radek
      sloupce.each_with_index do |s, i|
        pismeno = XLSSestavy.sloupec_pismeno x+i
        formule = case s.souctovy_radek
                    when :soucet; "SUBTOTAL(9,#{pismeno}#{rozsah.min}:#{pismeno}#{rozsah.max})"
                    when :pocet; "SUBTOTAL(9,#{pismeno}#{rozsah.min}:#{pismeno}#{rozsah.max})"
                    else
                      @ws.write y, x+i, '', format
                      next
                  end
        @ws.write_formula y, x+i, formule, format
      end
    end

  end
end