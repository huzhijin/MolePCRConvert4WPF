using System;

namespace MolePCRConvert4WPF.Core.Interfaces
{
    /// <summary>
    /// 导航服务接口
    /// </summary>
    public interface INavigationService
    {
        /// <summary>
        /// 导航到指定视图模型
        /// </summary>
        /// <typeparam name="TViewModel">视图模型类型</typeparam>
        void NavigateTo<TViewModel>() where TViewModel : class;
        
        /// <summary>
        /// 获取指定类型的视图模型实例
        /// </summary>
        /// <typeparam name="TViewModel">视图模型类型</typeparam>
        /// <returns>视图模型实例</returns>
        TViewModel? GetViewModel<TViewModel>() where TViewModel : class;
    }
} 