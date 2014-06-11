# encoding: utf-8
module XLSSestavy
  class Sloupec

    attr_accessor :tabulka

    def initialize(zahlavi, args={}, &block)
      @args = args.merge! zahlavi: zahlavi, fce: (args.delete(:fce) || block)
    end

    def zahlavi
      @zahlavi ||= I18n.translate(@args[:zahlavi], argumenty_hash.merge!(default: @args[:zahlavi]))
    end

    def argumenty_hash
      raise 'Není přiřazena tabulka.' unless @tabulka
      argumenty_sestavy = @tabulka.sestava.argumenty
      (@args[:argumenty] || []).inject({}) do |hash, arg_sym|
        hash[arg_sym] = argumenty_sestavy.send arg_sym
        hash
      end
    end

    def argumenty_array
      return @argumenty_array if defined? @argumenty_array
      arg_hash = argumenty_hash
      @argumenty_array = (@args[:argumenty] || []).map{|arg_sym| arg_hash[arg_sym]}
    end

    def fce
      @fce ||= @args[:fce]
    end

    def hodnota_pro(objekt)
      if fce.class==Proc
        fce.call(objekt, *argumenty_array)
      elsif objekt.respond_to? fce
        objekt.send(fce, *argumenty_array)
      elsif nil
        nil
      else
        raise "Není definováno pro @fce=#{fce.class}:#{fce}"
      end
    end

  end
end