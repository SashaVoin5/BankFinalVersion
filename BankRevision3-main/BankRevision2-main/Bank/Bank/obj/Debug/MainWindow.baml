   M S B A M L   `   `   `  ?????  ;Bank, Version=1.0.0.0, Culture=neutral, PublicKeyToken=nullclr-namespace:BankBank  5   	       Bank.MainWindow   5   	   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   $localclr-namespace:Bank     5   	     ? Title$  
MainWindow??5   	   $	??450??6   $	??800??6)   .??5	      ? $
??Left=?6   $	??450??6&   $	??Top=?63   $	??800??6K   .+?5   
   ?? $3??-C:\Users\Admin\Desktop\Модули\Logo1.JPG??6   $??719,10,20,371q?6A   ?? 5   
   -   5         ??Name$ btn_earn??5      $%??Рассчитать доход??6#   $??189,340,266,0q?6>   $	??Top=?6U   ???	#FF2C7DE0
???????$??30??6?   $??16(?6?   $??2??6?   $??2??6?    J? Effect 5      3  -System.Windows.Media.Effects.DropShadowEffect  5      ??5      ?? 5      5      ?? 5   
   $/??)Деньги в каждый карман??6   $
??Left=?61   $??
519,23,0,0q?6L   $	??Top=?6`   $??16(?6x   ???	#FF002060$??31??6?   $	??195??6?   ?? 5   
   $
??Left=?6   $	??200??6/   $??
121,94,0,0q?6<   $	??Top=?6P   $	??135??6h   ???	#FFDEEBF7 5        5         Color$ #deebf7??6#      	Direction$ 0Q?63      ShadowDepth$ 0Q?6A      
BlurRadius$	 100Q?6Q   5      ?? 5   
   $??Center=?6   $	??200??61   $??0,94,0,0q?6>   $	??Top=?6P   $	??128??6h   ???	#FFDEEBF7 5        5      $ #deebf7??6#   $ 0Q?63   $ 0Q?6A   $	 100Q?6Q   5      ?? 5   
   $
??Left=?6   $	??200??6/   $??
542,94,0,0q?6<   $	??Top=?6P   $	??129??6h   ???	#FFDEEBF7 5         5!      $ #deebf7??6#   $ 0Q?63   $ 0Q?6A   $	 100Q?6Q   5#      ?? 5$   
   $??годовых??6   $
??Left=?6"   $??167,251,0,0q?6=   $	??Top=?6R   ???	#FF002060$??18(?6?    J? RenderTransformOrigin$ 
0.65,0.793&?6?   ?? 5%   
   $??Пополняемый
??6   $??Center=?60   $??	0,153,0,0q?6M   $	??Top=?6`   ?? 5&   
   $b??\  Без пополнения
       Без снятия
Без капитализации??6   $
??Left=?6a   $??131,153,0,0q?6|   $	??Top=?6?   ?? 5'   
   $2??,Есть возможность снятия??6   $??Center=?62   $??	0,170,0,0q?6O   $	??Top=?6b   $??10(?6z   ?? 5(   
   $&?? С капитализацией
??6   $??Center=?60   $??	0,183,0,0q?6M   $	??Top=?6`   ?? 5)   
   $C??=Минимальный срок
      от 6 месяцев

??6   $??Center=?6L   $??	0,207,0,0q?6i   $	??Top=?6|   ?? 5*   
   $??годовых??6   $
??Left=?6"   $??377,251,0,0q?6=   $	??Top=?6R   ???	#FF002060$??18(?6?   $ 
0.65,0.793&?6?   ?? 5+   
   $??годовых??6   $
??Left=?6"   $??167,251,0,0q?6=   $	??Top=?6R   ???	#FF002060$??18(?6?   $ 
0.65,0.793&?6?   ?? 5,   
   $_??Y    Пополняемый
       Без снятия
Без капитализации??6   $
??Left=?6`   $??549,153,0,0q?6{   $	??Top=?6?   ?? 5-   
   $C??=Минимальный срок
      от 3 месяцев

??6   $
??Left=?6L   $??549,211,0,0q?6g   $	??Top=?6|   ?? 5.   
   $??годовых??6   $
??Left=?6"   $??584,251,0,0q?6=   $	??Top=?6R   ???	#FF002060$??18(?6?   $ 
0.65,0.793&?6?   ?? 5/   
   $??8%??6   $
??Left=?6   $??131,248,0,0q?68   $	??Top=?6M   $??22(?6e   ???	#FF0020606s   ?? 50   
   $%??Минимальный срок??6   $
??Left=?6+   $??128,204,0,0q?6F   $	??Top=?6[   $ 0.505,0.445&?6s   ?? 51   
   $??от 3 месяцев??6   $
??Left=?6'   $??148,222,0,0q?6B   $	??Top=?6W   ?? 52   
   $??Оптимальный??6   $??Center=?6&   $??	0,104,0,0q?6C   $	??Top=?6V   $??16(?6n   $
??Bold"?6|   ?? 53   
   $??Стандарт??6   $
??Left=?6#   $??568,104,0,0q?6>   $	??Top=?6S   $
??Bold"?6k   $??16(?6}   ?? 54   
   $??Стабильный??6   $
??Left=?6%   $??135,104,0,0q?6@   $	??Top=?6U   $??16(?6m   $
??Bold"?6{   ?? 55   
   $??5%??6   $
??Left=?6   $??342,247,0,0q?68   $	??Top=?6M   $??22(?6e   ???	#FF0020606s   ?? 56   
   $??6%??6   $
??Left=?6   $??549,247,0,0q?68   $	??Top=?6M   $??22(?6e   ???	#FF0020606s   5;      5<      