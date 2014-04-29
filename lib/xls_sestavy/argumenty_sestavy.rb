# encoding: utf-8
module XLSSestavy
  class ArgumentySestavy

    attr_accessor :od_data, :do_data, :uzivatel_id
    alias_method :k_datu, :od_data
    def k_datu=(datum); @od_data = datum; @do_data = datum; end

    def initialize(args={})
      args.each_pair{|k, v| send "#{k}=", v }
    end

    def jeden_datum?
      @od_data==@do_data
    end

    def argumenty_sloupce(arr)
      arr.map{|a| (a.class==Symbol && respond_to?(a)) ? send(a) : a}
    end

    def to_s
      return @to_s if defined? @to_s
      @to_s = if jeden_datum?
                "#{I18n.l(@od_data, format: :excel)[0..-2]}"
              else
                "#{I18n.l(@od_data, format: :excel)[0..-2]}_#{I18n.l(@do_data, format: :excel)[0..-2]}"
              end
    end

  end
end
