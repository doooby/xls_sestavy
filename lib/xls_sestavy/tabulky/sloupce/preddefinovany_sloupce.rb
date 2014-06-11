# encoding: utf-8
module XLSSestavy
  class PreddefinovanySloupec < Sloupec

    attr_reader :klic

    def initialize(klic, zahlavi, args={}, &block)
      super zahlavi, args, &block
      @klic = klic.to_sym
    end

  end
end