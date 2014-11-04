# encoding: utf-8
module XLSSestavy
  class PromenneRadku
    attr_reader :tabulka, :objekt

    def initialize
      @metody = {}
    end

    def nastav(klic, &block)
      @metody[klic] = block
    end

    def pro_tabulku(tabulka)
      @tabulka = tabulka
    end

    def pro_objekt(objekt)
      unless @objekt == objekt
        @hodnoty = {}
        @objekt = objekt
      end
      self
    end

    def [](klic)
      h = @hodnoty[klic]
      unless h
        m = @metody[klic]
        h = m ? m.call(@objekt) : nil
        @hodnoty[klic] = h
      end
      h
    end

  end
end