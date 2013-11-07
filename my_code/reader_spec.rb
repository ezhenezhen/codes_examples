# -*- encoding : utf-8 -*-
require 'spec_helper'

describe FeedReaders::ArtMaster do
  before(:each) do
    @owner = FactoryGirl.create(:owner, id: FeedReaders::ArtMaster.owner_id)

    # construction 1 with 2 sides
    @board1 = FactoryGirl.create(:construction, owner: @owner, code_from_owner: '1', type_id: Construction::BILBOARD_3X6)
    @s11 = FactoryGirl.create(:surface, name: 'А', construction: @board1)
    @s12 = FactoryGirl.create(:surface, name: 'Б', construction: @board1)

    # construction 2 with 2 sides
    @board2 = FactoryGirl.create(:construction, owner: @owner, code_from_owner: '2', type_id: Construction::BILBOARD_3X6)
    @s21 = FactoryGirl.create(:surface, name: 'А', construction: @board2)

    # construction in DB but not included in the feed, should be set as sold
    @board3 = FactoryGirl.create(:construction, owner: @owner, code_from_owner: '5', type_id: Construction::BILBOARD_3X6)
    @s31 = FactoryGirl.create(:surface, construction: @board3)
  end

  describe 'regular feed' do
    before(:all) { Timecop.travel(Time.local(2013)) }
    after(:all)  { Timecop.return }

    before(:each) do
      @feed = File.join(Rails.root, %w(spec fixtures xls_feeds art_master.xls))
      @expected = {
          @s11.id => {
              'M0913' => { status: Stock::SOLD,      base: 18_000, sell: 18_000 },
              'M1013' => { status: Stock::SOLD,      base: 18_000, sell: 18_000 },
              'M1113' => { status: Stock::SOLD,      base: 18_000, sell: 18_000 },
              'M1213' => { status: Stock::RESERVED,  base: 18_000, sell: 18_000 },
          },
          @s12.id => {
              'M0913' => { status: Stock::SOLD,      base: 18_000, sell: 18_000 },
              'M1013' => { status: Stock::SOLD,      base: 18_000, sell: 18_000 },
              'M1113' => { status: Stock::SOLD,      base: 18_000, sell: 18_000 },
              'M1213' => { status: Stock::SOLD,      base: 18_000, sell: 18_000 },
          },
          @s21.id => {
              'M0913' => { status: Stock::SOLD,      base: 18_000, sell: 18_000 },
              'M1013' => { status: Stock::SOLD,      base: 18_000, sell: 18_000 },
              'M1113' => { status: Stock::RESERVED,  base: 18_000, sell: 18_000 },
              'M1213' => { status: Stock::FREE,      base: 18_000, sell: 18_000 },
          },
          @s31.id => {
              'M0913' => { status: Stock::SOLD },
              'M1013' => { status: Stock::SOLD },
              'M1113' => { status: Stock::SOLD },
              'M1213' => { status: Stock::SOLD },
          }
      }
    end

    it 'should read surfaces from different sheets' do
      got = FeedReaders::ArtMaster.new(@feed).do
      #FeedReader.diff_results(got, @expected)
      got.should == @expected
    end

    it 'should read selective' do
      @expected.delete(@s31.id)
      got = FeedReaders::ArtMaster.new(@feed, missed_surfaces_are_not_sold: true).do
      #FeedReader.diff_results(got, @expected)
      got.should == @expected
    end

    it 'should save unrecognized' do
      r = FeedReaders::ArtMaster.new(@feed)
      r.do
      r.unrecognized.should == [
          "Unknown [3] side X <a href='http://www.art-master-ufa.ru/image/3bb.jpg' target='_blank'>ссылка</a>"
      ]
    end
  end
end