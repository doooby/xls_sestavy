# encoding: utf-8
module XLSSestavy
  class Xls::TabulkaXls < Tabulka

    # ukotvit_zahlavi, format_zahlavi, format_souctu, souctove_radky, format_dat, velikost_sady, vyska_zahlavi
    def initialize(sestava, worksheet, args={}, &block)
      raise 'Sestava musí být typu SestavaXls' unless sestava.kind_of? SestavaXls
      @worksheet = worksheet
      super sestava, args, &block
    end

    def vypis_radek_zahlavi(pozice)
      r, c = XLSSestavy::Xls.ciselne_souradnice(pozice).first
      format_zahlavi = @sestava.get_format @args[:format_zahlavi] || :radek_zahlavi

      @sloupce.each_with_index do |s, i|
        format = if s.opt(:xls_format).class==Hash
                   format_hash = s.opt(:xls_format).clone
                   format_hash.delete :num_format
                   @sestava.alter_format format_zahlavi, format_hash
                 else
                   format_zahlavi
                 end
        @worksheet.write r, c+i, s.zahlavi, format
      end
    end

    def vypis_radek_souctu(pozice, rozsah=nil)
      r, c = XLSSestavy::Xls.ciselne_souradnice(pozice).first
      format_souctu = @sestava.get_format @args[:format_souctu] || :radek_souctu

      @sloupce.each_with_index do |s, i|
        format_hash = {num_format: XLSSestavy::Xls.num_format(s.opt(:datovy_typ))}
        format_hash.merge! s.opt(:xls_format) if s.opt(:xls_format).class == Hash
        format = @sestava.alter_format format_souctu, format_hash

        pismeno = XLSSestavy::Xls.sloupec_pismeno c + i
        case rozsah && s.opt(:radek_souctu)
          when :soucet
            formule = "SUBTOTAL(9,#{pismeno}#{rozsah.min}:#{pismeno}#{rozsah.max})"
            @worksheet.write_formula r, c+i, formule, format
          when :pocet
            formule = "SUBTOTAL(3,#{pismeno}#{rozsah.min}:#{pismeno}#{rozsah.max})"
            @worksheet.write_formula r, c+i, formule, format
          else
            @worksheet.write r, c+i, nil, format
        end
      end
    end

    def vypis_radky_dat(pozice, radky)
      r, c = XLSSestavy::Xls.ciselne_souradnice(pozice).first
      radku_vypsano = 0

      format_dat = @sestava.get_format @args[:format_dat] || :radky_dat
      formaty = @sloupce.map do |s|
        format_hash = {num_format: XLSSestavy::Xls.num_format(s.opt(:datovy_typ))}
        format_hash.merge! s.opt(:xls_format) if s.opt(:xls_format).class == Hash
        @sestava.alter_format format_dat, format_hash
      end

      if radky.class == ActiveRecord::Relation
        velikost_sady = @args[:velikost_sady] || 200
        radky.find_in_batches batch_size: velikost_sady do |sada|
          aktualni_pozice = "#{XLSSestavy::Xls.sloupec_pismeno c}#{r + 1 + radku_vypsano}"
          vypis_blok_dat aktualni_pozice, sada, formaty
          radku_vypsano += sada.length
        end
      else
        vypis_blok_dat pozice, radky, formaty
        radku_vypsano = radky.length
      end

      radku_vypsano
    end

    def vypis_blok_dat(pozice, radky, formaty=[])
      r, c = XLSSestavy::Xls.ciselne_souradnice(pozice).first

      radky.each do |objekt|
        @sloupce.each_with_index do |s, i|
          hodnota = XLSSestavy::Xls.douprav_hodnotu_bunky s.hodnota_pro(objekt)
          case s.opt(:datovy_typ)
            when :cas, :datum
              @worksheet.write_date_time r, c+i, hodnota, formaty[i]
            else
              @worksheet.write r, c+i, hodnota, formaty[i]
          end
        end
        r += 1
      end
    end

    def vypis(pozice, radky)
      r, c = XLSSestavy::Xls.ciselne_souradnice(pozice).first
      sloupec_pismeno = XLSSestavy::Xls.sloupec_pismeno c
      puvodni_r = r
      pocet_radku = radky.count

      soucty_od = puvodni_r + 1
      souctove_radky = @args[:souctove_radky]
      souctove_radky = :prazdne if pocet_radku==0 && souctove_radky

      #šířky sloupců
      @sloupce.each_with_index do |s, i|
        sirka = s.opt(:sirka_sloupce)
        next unless sirka
        si = c + i
        @worksheet.set_column si, si, XLSSestavy::Xls.col_cm_to_p(sirka)
      end

      #hlavičky
      vypis_radek_zahlavi pozice
      @worksheet.set_row r, XLSSestavy::Xls.row_cm_to_p(@args[:vyska_zahlavi]  || 1.3)
      r += 1

      #součtové řádky nad
      if souctove_radky==:nad || souctove_radky==:nad_pod
        soucty_od += 1
        vypis_radek_souctu "#{sloupec_pismeno}#{r + 1}", (soucty_od..soucty_od+pocet_radku-1)
        r += 1
      end

      #ukotvení
      @worksheet.freeze_panes r, 0 if @args[:ukotvit_zahlavi]

      #vypsání dat
      r += vypis_radky_dat("#{sloupec_pismeno}#{r + 1}", radky)

      #součty pod
      if souctove_radky==:pod || souctove_radky==:nad_pod
        vypis_radek_souctu "#{sloupec_pismeno}#{r + 1}", (soucty_od..soucty_od+pocet_radku-1)
        r += 1
      end

      #prázdný součtový řádek
      if souctove_radky==:prazdne
        vypis_radek_souctu "#{sloupec_pismeno}#{r + 1}"
        r += 1
      end

      r - puvodni_r
    end

  end
end