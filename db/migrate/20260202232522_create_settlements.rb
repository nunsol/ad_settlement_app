class CreateSettlements < ActiveRecord::Migration[8.1]
  def change
    create_table :settlements do |t|
      t.string :period
      t.string :status
      t.integer :total_rows
      t.integer :matched_rows
      t.integer :unmatched_rows
      t.json :unmatched_accounts
      t.json :agency_distribution
      t.text :error_message

      t.timestamps
    end
    add_index :settlements, :status
  end
end
