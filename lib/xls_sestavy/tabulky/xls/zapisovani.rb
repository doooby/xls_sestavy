# encoding: utf-8
module XLSSestavy
  module Xls
    module Zapisovani

      # roztahuje se na (defaultně) 24 sloupců v prvním řádku
      def sestava_nadpis(text, roztahnout=12, radek=0)
        @ws.merge_range radek, 0, radek, roztahnout-1, text, get_format(:sestava_nadpis)
        @ws.set_row radek, XLSSestavy::Xls.row_cm_to_p(1)
      end

      # pozice ve standartních excel souřadnicích 'A2', 'B3:B5' (spojení buňek)
      # vyska znamená výška daného řádku v cm
      def sestava_napdis2(text, pozice, vyska = 0.7)
        zapis pozice, text, get_format(:sestava_nadpis2)
        @ws.set_row ciselne_souradnice(pozice).first.first, XLSSestavy::Xls.row_cm_to_p(vyska)
      end

      #pozice ve standartních excel souřadnicích 'A2', 'B3:B5' (spojení buňek)
      def sestava_cas_vytvoreni(pozice='A2')
        cas = "začátek zpracování: #{I18n.localize Time.zone.now}"
        zapis pozice, cas, get_format(:default)
      end

      # pozice bunkdy jsou ve standartních souřadnicích ( 'A3')
      def zapis(pozice, hodnota, format=nil)
        format = get_format unless format
        z_pole, do_pole = XLSSestavy::Xls.ciselne_souradnice pozice
        if do_pole
          @ws.merge_range z_pole[0], z_pole[1], do_pole[0], do_pole[1], hodnota, format
        else
          @ws.write z_pole[0], z_pole[1], hodnota, format
        end
      end

      def zapis_radu(pozice, hodnoty, format=nil)
        format = get_format unless format
        r, c = XLSSestavy::Xls.ciselne_souradnice(pozice).first
        hodnoty.each do |h|
          @ws.write r, c, h, format
          c += 1
        end
      end

      def zapis_sloupec(pozice, hodnoty, format=nil)
        format = get_format unless format
        r, c = XLSSestavy::Xls.ciselne_souradnice(pozice).first
        hodnoty.each do |h|
          @ws.write r, c, h, format
          r += 1
        end
      end

      def zapis_tabulku(pozice, radky, args={}, &block)
        tabulka = XLSSestavy::TabulkaXls.new self, @ws, args, &block
        tabulka.vypis pozice, radky
      end

    end
  end
end