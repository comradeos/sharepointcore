import type { IProductSource } from '../../services/products/IProductSource';

/** Описує заголовок каталогу та джерело даних незалежного компонента товарів. */
export interface IProductsProps {
  description: string;
  source: IProductSource;
}
