export interface IProduct {
  id: number;
  title: string;
  description: string;
  category: string;
  price: number;
  thumbnail: string;
}

export interface IProductPage {
  products: IProduct[];
  total: number;
}

/** Кожне джерело перетворює відповідь на спільну модель товарів. */
export interface IProductSource {
  readonly id: string;
  readonly name: string;
  /** Повертає сторінку товарів із можливістю скасування запиту. */
  getProducts(skip: number, limit: number, signal: AbortSignal): Promise<IProductPage>;
}
