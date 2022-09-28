import { JpStockGetter } from "./JpStockGetter"

/**
 * 株価情報の更新 
 * 
 * @date 2022/9/28
 */
export function Update() 
{
    let instance = new JpStockGetter();
    instance.GetJpStock();
}