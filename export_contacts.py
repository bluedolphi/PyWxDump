# -*- coding: utf-8 -*-
"""
微信好友和群信息综合导出工具
基于PyWxDump库同时导出好友和群信息
"""
import json
import os
import csv
from typing import Dict, List, Any
from datetime import datetime

import pywxdump
from pywxdump import MicroHandler, get_wx_info, get_wx_db

from export_friends import WxFriendsExporter
from export_groups import WXGroupsExporter


class WXContactsExporter:
    """微信联系人综合导出器（好友+群）"""
    
    def __init__(self, db_path: str = None, key: str = None):
        """
        初始化综合导出器
        
        Args:
            db_path: 微信数据库路径
            key: 数据库解密密钥
        """
        self.db_path = db_path
        self.key = key
        self.friends_exporter = WxFriendsExporter(db_path, key)
        self.groups_exporter = WXGroupsExporter(db_path, key)
        
    def export_all_contacts(self) -> Dict[str, Any]:
        """
        导出所有联系人信息（好友+群）
        
        Returns:
            包含所有联系人信息的字典
        """
        print("开始导出微信联系人信息...")
        
        # 导出好友信息
        print("正在获取好友信息...")
        friends = self.friends_exporter.get_friends_info()
        
        # 导出群信息
        print("正在获取群信息...")
        groups = self.groups_exporter.get_groups_info()
        
        # 统计信息
        total_contacts = len(friends) + len(groups)
        export_time = datetime.now().isoformat()
        
        export_data = {
            "export_info": {
                "export_time": export_time,
                "total_contacts": total_contacts,
                "friends_count": len(friends),
                "groups_count": len(groups)
            },
            "friends": friends,
            "groups": groups,
            "statistics": self._generate_statistics(friends, groups)
        }
        
        print(f"导出完成! 共 {total_contacts} 个联系人")
        return export_data
    
    def _generate_statistics(self, friends: List[Dict], groups: List[Dict]) -> Dict[str, Any]:
        """生成统计信息"""
        stats = {
            "friends": {},
            "groups": {},
            "combined": {}
        }
        
        # 好友统计
        if friends:
            gender_count = {0: 0, 1: 0, 2: 0}
            for friend in friends:
                gender_count[friend['gender']] += 1
            
            stats["friends"] = {
                "total": len(friends),
                "gender": {
                    "male": gender_count[1],
                    "female": gender_count[2],
                    "unknown": gender_count[0]
                },
                "with_remark": len([f for f in friends if f['remark']]),
                "with_wechat_id": len([f for f in friends if f['wechat_id']])
            }
        
        # 群统计
        if groups:
            total_members = sum(g['member_count'] for g in groups)
            stats["groups"] = {
                "total": len(groups),
                "total_members": total_members,
                "avg_members_per_group": total_members / len(groups) if groups else 0,
                "largest_group": max(groups, key=lambda x: x['member_count']) if groups else None,
                "smallest_group": min(groups, key=lambda x: x['member_count']) if groups else None
            }
        
        # 综合统计
        stats["combined"] = {
            "total_contacts": len(friends) + len(groups),
            "total_interactions": len(friends) + sum(g['member_count'] for g in groups)
        }
        
        return stats
    
    def export_to_json(self, output_path: str = None) -> bool:
        """
        导出到JSON文件
        
        Args:
            output_path: 输出文件路径
            
        Returns:
            导出是否成功
        """
        try:
            data = self.export_all_contacts()
            
            if not output_path:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                output_path = f"out/wx_contacts_{timestamp}.json"
            else:
                # 确保路径在out目录下
                if not output_path.startswith('out/'):
                    output_path = f"out/{output_path}"
            
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            
            print(f"联系人信息已导出到: {output_path}")
            return True
            
        except Exception as e:
            print(f"JSON导出失败: {e}")
            return False
    
    def export_to_csv(self, output_path: str = None) -> bool:
        """
        导出到CSV文件（分别导出好友和群）
        
        Args:
            output_path: 输出文件路径（不带扩展名）
            
        Returns:
            导出是否成功
        """
        try:
            if not output_path:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                base_path = f"out/wx_contacts_{timestamp}"
            else:
                # 确保路径在out目录下
                if not output_path.startswith('out/'):
                    base_path = f"out/{output_path}"
                else:
                    base_path = output_path
            
            # 导出好友CSV
            friends_csv = f"{base_path}_friends.csv"
            friends_success = self.friends_exporter.export_to_csv(friends_csv)
            
            # 导出群CSV
            groups_csv = f"{base_path}_groups.csv"
            groups_success = self.groups_exporter.export_to_csv(groups_csv)
            
            if friends_success and groups_success:
                print(f"联系人信息已导出到:")
                print(f"   好友: {friends_csv}")
                print(f"   群聊: {groups_csv}")
                return True
            else:
                return False
            
        except Exception as e:
            print(f"CSV导出失败: {e}")
            return False
    
    def export_to_html(self, output_path: str = None) -> bool:
        """
        导出到HTML文件（可视化报告）
        
        Args:
            output_path: 输出文件路径
            
        Returns:
            导出是否成功
        """
        try:
            data = self.export_all_contacts()
            
            if not output_path:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                output_path = f"out/wx_contacts_{timestamp}.html"
            else:
                # 确保路径在out目录下
                if not output_path.startswith('out/'):
                    output_path = f"out/{output_path}"
            
            # 生成HTML报告
            html_content = self._generate_html_report(data)
            
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(html_content)
            
            print(f"HTML报告已导出到: {output_path}")
            return True
            
        except Exception as e:
            print(f"HTML导出失败: {e}")
            return False
    
    def _generate_html_report(self, data: Dict[str, Any]) -> str:
        """生成HTML报告"""
        stats = data['statistics']
        friends = data['friends']
        groups = data['groups']
        
        html = f"""
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>微信联系人报告</title>
    <style>
        body {{ font-family: 'Microsoft YaHei', Arial, sans-serif; margin: 20px; background-color: #f5f5f5; }}
        .container {{ max-width: 1200px; margin: 0 auto; background: white; padding: 20px; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }}
        .header {{ text-align: center; color: #07C160; margin-bottom: 30px; }}
        .section {{ margin-bottom: 30px; }}
        .stats {{ display: flex; justify-content: space-around; flex-wrap: wrap; margin: 20px 0; }}
        .stat-card {{ background: #f8f9fa; padding: 20px; border-radius: 8px; text-align: center; min-width: 150px; margin: 10px; }}
        .stat-number {{ font-size: 24px; font-weight: bold; color: #07C160; }}
        .stat-label {{ color: #666; margin-top: 5px; }}
        .contacts-table {{ width: 100%; border-collapse: collapse; margin-top: 10px; }}
        .contacts-table th, .contacts-table td {{ border: 1px solid #ddd; padding: 12px; text-align: left; }}
        .contacts-table th {{ background-color: #07C160; color: white; }}
        .contacts-table tr:nth-child(even) {{ background-color: #f9f9f9; }}
        .section-title {{ color: #07C160; border-bottom: 2px solid #07C160; padding-bottom: 5px; }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>微信联系人报告</h1>
            <p>导出时间: {data['export_info']['export_time']}</p>
        </div>
        
        <div class="section">
            <h2 class="section-title">总体统计</h2>
            <div class="stats">
                <div class="stat-card">
                    <div class="stat-number">{data['export_info']['total_contacts']}</div>
                    <div class="stat-label">总联系人</div>
                </div>
                <div class="stat-card">
                    <div class="stat-number">{data['export_info']['friends_count']}</div>
                    <div class="stat-label">好友</div>
                </div>
                <div class="stat-card">
                    <div class="stat-number">{data['export_info']['groups_count']}</div>
                    <div class="stat-label">群聊</div>
                </div>
            </div>
        </div>
        
        <div class="section">
            <h2 class="section-title">好友统计</h2>
            <div class="stats">
                <div class="stat-card">
                    <div class="stat-number">{stats['friends']['gender']['male']}</div>
                    <div class="stat-label">男性好友</div>
                </div>
                <div class="stat-card">
                    <div class="stat-number">{stats['friends']['gender']['female']}</div>
                    <div class="stat-label">女性好友</div>
                </div>
                <div class="stat-card">
                    <div class="stat-number">{stats['friends']['with_remark']}</div>
                    <div class="stat-label">有备注</div>
                </div>
            </div>
        </div>
        
        <div class="section">
            <h2 class="section-title">群聊统计</h2>
            <div class="stats">
                <div class="stat-card">
                    <div class="stat-number">{stats['groups']['total_members']}</div>
                    <div class="stat-label">群成员总数</div>
                </div>
                <div class="stat-card">
                    <div class="stat-number">{stats['groups']['avg_members_per_group']:.1f}</div>
                    <div class="stat-label">平均群成员</div>
                </div>
            </div>
        </div>
        
        <div class="section">
            <h2 class="section-title">好友列表</h2>
            <table class="contacts-table">
                <thead>
                    <tr>
                        <th>昵称</th>
                        <th>备注</th>
                        <th>微信号</th>
                        <th>性别</th>
                    </tr>
                </thead>
                <tbody>
        """
        
        gender_map = {0: '未知', 1: '男', 2: '女'}
        for friend in friends[:50]:  # 只显示前50个
            html += f"""
                    <tr>
                        <td>{friend['nickname'][:20]}</td>
                        <td>{friend['remark'][:20]}</td>
                        <td>{friend['wechat_id'][:20]}</td>
                        <td>{gender_map.get(friend['gender'], '未知')}</td>
                    </tr>
            """
        
        if len(friends) > 50:
            html += f"""
                    <tr>
                        <td colspan="4" style="text-align: center; color: #666;">
                            ... 还有 {len(friends) - 50} 个好友
                        </td>
                    </tr>
            """
        
        html += """
                </tbody>
            </table>
        </div>
        
        <div class="section">
            <h2 class="section-title">群聊列表</h2>
            <table class="contacts-table">
                <thead>
                    <tr>
                        <th>群名称</th>
                        <th>成员数</th>
                        <th>群状态</th>
                    </tr>
                </thead>
                <tbody>
        """
        
        for group in groups[:30]:  # 只显示前30个群
            html += f"""
                    <tr>
                        <td>{group['display_name'][:30]}</td>
                        <td>{group['member_count']}</td>
                        <td>{group.get('chatroom_flag', '')}</td>
                    </tr>
            """
        
        if len(groups) > 30:
            html += f"""
                    <tr>
                        <td colspan="3" style="text-align: center; color: #666;">
                            ... 还有 {len(groups) - 30} 个群聊
                        </td>
                    </tr>
            """
        
        html += """
                </tbody>
            </table>
        </div>
    </div>
</body>
</html>
        """
        
        return html
    
    def print_summary(self):
        """打印综合摘要"""
        print("\n" + "="*60)
        print("微信联系人综合摘要")
        print("="*60)
        
        # 导出好友摘要
        self.friends_exporter.print_friends_summary()
        
        print("\n" + "-"*60)
        
        # 导出群摘要
        self.groups_exporter.print_groups_summary()
        
        print("="*60)


def main():
    """主函数"""
    print("微信好友和群信息综合导出工具")
    print("=" * 50)
    
    try:
        # 创建综合导出器实例
        exporter = WXContactsExporter()
        
        # 打印综合摘要
        exporter.print_summary()
        
        # 导出到不同格式
        print("\n正在导出文件...")
        
        # 导出到JSON
        if exporter.export_to_json():
            print("JSON导出成功")
        
        # 导出到CSV
        if exporter.export_to_csv():
            print("CSV导出成功")
        
        # 导出到HTML
        if exporter.export_to_html():
            print("HTML报告导出成功")
        
        print("\n联系人信息导出完成!")
        
    except Exception as e:
        print(f"程序执行失败: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()