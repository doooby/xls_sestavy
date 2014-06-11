# encoding: utf-8
module XLSSestavy
  module Xls

    #převod cm na points pro řádky
    def self.row_cm_to_p(cm)
      (cm*28.3464567).to_i
    end

    #převod cm na points pro sloupce
    #přibližně .. (nenalezen přesný výpočet)
    def self.col_cm_to_p(cm)
      (cm*28.3464567/5.6).to_i
    end

    #převod z mm na palce
    def self.mm_to_inch(mm)
      mm/25.4
    end

    #číslo sloupce ze znaků ('A' => 0)
    # do maximální hodnoty 'ZZ'
    def self.sloupec_cislo(znak)
      znak.upcase!
      if znak.length==1
        znak.ord-65
      elsif znak.length==2
        (znak[0].ord-64)*26 + znak[1].ord-65
      end
    end

    #písmeno sloupce z čísla (0 => 'A')
    # do maximální hodnoty 'ZZ'
    def self.sloupec_pismeno(cislo)
      return (cislo+65).chr if cislo < 26
      a = (cislo/26) - 1
      b = cislo%26
      "#{(a+65).chr}#{(b+65).chr}"
    end

    #převod standartních souřadnic na pole číselných souradnic:
    # 'A3' => [[2,0]], 'A5:B7' => [[4,0],[6,1]] <radek, sloupec>
    def self.ciselne_souradnice(bunky)
      a = bunky.match /^(\D+)(\d+)(:(\D+)(\d+))?$/
      return unless a
      souradnice = [[a[2].to_i-1, sloupec_cislo(a[1])]]
      souradnice << [a[5].to_i-1, sloupec_cislo(a[4])] if a[3]
      souradnice
    end

    #doupravuje hodnotu buňky, aby nedošlo na předvídatelné konflikty.
    #Důležité pro všechny typy času/datumu, protože to je potřeba převést na textový řetězec pro excel stravitelný
    def self.douprav_hodnotu_bunky(hodnota)
      return '' unless hodnota
      case hodnota
        when Hash
          YAML.dump hodnota
        when Array
          hodnota.join ', '
        when Time, DateTime, Date
          I18n.localize hodnota, format: :excel
        else
          hodnota
      end
    end

    #klíč num_formátu pro excel
    def self.num_format(sloupec)
      sym = sloupec.arg(:datovy_typ)
      sym && case sym
               when :cas; 'yyyy-MM-dd HH:mm:ss'
               when :datum; 'd. M. yyyy'
               when :suma; "#,###0.00 #{@def_mena}"
               when :pocet; '#,##0'
               when :cislo; '#0'
               else raise "nedefinovaný num_format: #{sym}"
             end
    end

    def self.def_mena=(mena)
      @def_mena = case mena
                    when '€';  '[$€-4B1]'
                    when 'Kč'; '[$Kč-405]'
                    else;      '??'
                  end
    end
    self.def_mena = 'Kč'

  end
end