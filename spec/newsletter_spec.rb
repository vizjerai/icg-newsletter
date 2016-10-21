require 'spec_helper'
require_relative '../lib/newsletter'

describe Newsletter do
  describe '#past_due?(date)' do
    let(:newsletter) { Newsletter.new('') }

    [
      Date.new(2016, 9, 30),
      Date.new(2016, 10, 31),
      Date.new(2017, 3, 31)
    ].each do |date|
      it "#{date} to not be past due" do
        Timecop.freeze(Date.new(2016, 10, 21)) do
          expect(newsletter.send(:past_due?, date)).to be_falsey
        end
      end
    end

    [
      Date.new(2016, 9, 29),
      Date.new(2016, 8, 31)
    ].each do |date|
      it "#{date} to be past due" do
        Timecop.freeze(Date.new(2016, 10, 21)) do
          expect(newsletter.send(:past_due?, date)).to be_truthy
        end
      end
    end

  end
end
