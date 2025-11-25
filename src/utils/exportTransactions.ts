import * as XLSX from 'xlsx';
import { Transaction } from '@/context/BudgetContext';

export const exportTransactionsByMonth = (transactions: Transaction[]) => {
  if (transactions.length === 0) {
    return;
  }

  // Group transactions by year-month
  const groupedByMonth = transactions.reduce((acc, transaction) => {
    const date = new Date(transaction.date);
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const key = `${year}-${month}`;
    
    if (!acc[key]) {
      acc[key] = [];
    }
    acc[key].push(transaction);
    return acc;
  }, {} as Record<string, Transaction[]>);

  // Create a workbook for each month
  Object.entries(groupedByMonth).forEach(([monthKey, monthTransactions]) => {
    // Sort transactions by date
    const sortedTransactions = monthTransactions.sort((a, b) => 
      new Date(a.date).getTime() - new Date(b.date).getTime()
    );

    // Format data for Excel
    const data = sortedTransactions.map(transaction => {
      const date = new Date(transaction.date);
      const year = date.getFullYear();
      const month = String(date.getMonth() + 1).padStart(2, '0');
      const day = String(date.getDate()).padStart(2, '0');
      
      return {
        Date: `${year}-${month}-${day}`,
        Description: transaction.description,
        Type: transaction.type === 'income' ? 'Income' : 'Expense',
        Amount: transaction.amount
      };
    });

    // Create worksheet and workbook
    const worksheet = XLSX.utils.json_to_sheet(data);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, monthKey);

    // Generate and download file
    const fileName = `transactions-${monthKey}.xlsx`;
    XLSX.writeFile(workbook, fileName);
  });
};
