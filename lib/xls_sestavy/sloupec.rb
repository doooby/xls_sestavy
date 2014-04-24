# encoding: utf-8
module XLSSestavy
  class Sloupec

    attr_accessor :zahlavi, :sirka, :format, :num_format, :souctovy_radek

    def initialize(zahlavi, args={}, &block)
      @zahlavi= zahlavi
      @fce = args[:fce] || block
      nastav args
    end

    def nastav(args={})
      @parametry = args[:parametry] || []
      @sirka = args[:sirka] || 3
      @souctovy_radek = args[:souctovy_radek]
      @format = args[:format]
      @num_format = args[:num_format]
      if @num_format
        @format ||= {}
        @format[:num_format] = XLSSestavy.num_format @num_format
      end
    end

    def nastav_argumenty(argumenty)
      @argumenty_array = argumenty.argumenty_sloupce(@parametry)
      @zahlavi.gsub!(/%\w+%/){|m| argumenty.send m[1..-2]}
    end

    def hodnota_pro(objekt)
      if @fce.class==Proc
        @fce.call(objekt, *@argumenty_array)
      elsif objekt.respond_to? @fce
        objekt.send(@fce, *@argumenty_array)
      else
        raise "Není definováno pro @fce=#{@fce.class}:#{@fce}"
      end
    end

  end
end