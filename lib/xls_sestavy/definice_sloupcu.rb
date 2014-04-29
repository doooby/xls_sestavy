# encoding: utf-8
module XLSSestavy
  class DefiniceSloupcu < RadaSloupcu

    def self.[](klic)
      DefiniceSloupcu.definice(self)[klic]
    end

    def self.pridej_jako_definice
      DefiniceSloupcu.cache_definic[self] = self
    end

    def definice; end

    def definuj(*args, &block)
      @sloupce << PreddefinovanySloupec.new(*args, &block)
    end

    private

    def self.definice(klass)
      ds = DefiniceSloupcu.cache_definic[klass]
      if ds.class==Class
        ds = ds.new
        ds.definice
        ds.aktualizuj_seznam
        @cache_definic[klass] = ds
      end
      ds
    end

    def self.cache_definic
      @cache_definic ||= {}
    end

  end
end