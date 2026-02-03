class CreateSettlementFiles < ActiveRecord::Migration[8.1]
  def change
    create_table :settlement_files do |t|
      t.references :settlement, null: false, foreign_key: true
      t.string :file_type
      t.string :original_filename
      t.string :status
      t.integer :rows_count
      t.integer :matched_count

      t.timestamps
    end
  end
end
