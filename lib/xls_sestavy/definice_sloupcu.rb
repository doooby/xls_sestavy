# encoding: utf-8
module XLSSestavy
  class DefiniceSloupcu < RadaSloupcu

    def self.[](klic)
      ds = DefiniceSloupcu.cache_definic[klic]
      if ds.class==Class
        ds = ds.new
        ds.definice
        ds.aktualizuj_seznam
        DefiniceSloupcu.cache_definic[klic] = ds
      end
      ds
    end

    def self.pridej_jako_definice(klic)
      DefiniceSloupcu.cache_definic[klic] = self
    end

    def definice; end

    def definuj(*args, &block)
      @sloupce << PreddefinovanySloupec.new(*args, &block)
    end

    private

    def self.cache_definic
      @cache_definic ||= {}
    end

  end
end